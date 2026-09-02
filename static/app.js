// ============================================================================
// Общее: переключение вкладок
// ============================================================================
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
    if (btn.dataset.tab === "tenders") loadTenders();
    if (btn.dataset.tab === "borderline") loadBorderline();
    if (btn.dataset.tab === "favorites") loadFavorites();
    if (btn.dataset.tab === "keywords") loadKeywords();
    if (btn.dataset.tab === "ai") loadAiSettings();
    if (btn.dataset.tab === "ai-rescan") pollAiRescanState();
  });
});

async function api(path, opts) {
  const res = await fetch(path, opts);
  return res.json();
}

// ============================================================================
// Вкладка "Сбор данных"
// ============================================================================
let lastLogId = 0;

function fmtState(state) {
  const labels = {
    idle: "не запущено", launching: "открываю Edge...", waiting_tab: "жду вкладку",
    ready: "готово к запуску", running: "идёт сбор", stopping: "останавливаюсь...",
    stopped: "остановлено", error: "ошибка",
  };
  return labels[state] || state;
}

function renderRunState(s) {
  const badge = document.getElementById("state-badge");
  badge.textContent = fmtState(s.state);
  badge.className = "badge " + s.state;

  const baseUrlEl = document.getElementById("base-url-line");
  baseUrlEl.textContent = s.base_url ? ("URL: " + s.base_url) : "";
  baseUrlEl.title = s.base_url || "";

  const mode = s.filter_mode || "keywords";
  document.querySelectorAll('input[name="collection-filter-mode"]').forEach(radio => {
    if (document.activeElement !== radio) radio.checked = radio.value === mode;
  });

  const urlInput = document.getElementById("search-url-input");
  if (document.activeElement !== urlInput && s.base_url && !urlInput.value) {
    urlInput.value = s.base_url;
  }
  urlInput.disabled = s.state === "running";

  const lastLine = document.getElementById("last-finished-line");
  if (s.restart_pending) {
    lastLine.textContent = "Сбор прервался с ошибкой — автоматический перезапуск запланирован через ~30 сек. Нажмите «Остановить», чтобы отменить.";
  } else if (s.state === "running" || s.state === "stopping") {
    lastLine.textContent = "Сбор идёт сейчас...";
  } else if (s.last_finished_at) {
    lastLine.textContent = "Статистика последнего сбора (завершён: " + s.last_finished_at + "):";
  } else {
    lastLine.textContent = "Сбор ещё ни разу не запускался.";
  }

  const autoRestartCheckbox = document.getElementById("auto-restart-checkbox");
  if (document.activeElement !== autoRestartCheckbox) {
    autoRestartCheckbox.checked = !!s.auto_restart;
  }
  const autoRestartInfo = document.getElementById("auto-restart-info");
  autoRestartInfo.textContent = s.auto_restart_count > 0 ? `Авто-перезапусков за сессию: ${s.auto_restart_count}` : "";

  setStatValue("s-page", s.counts.page);
  setStatValue("s-checked", s.counts.checked);
  setStatValue("s-found", s.counts.found_it);
  setStatValue("s-astana", s.counts.found_astana);
  setStatValue("s-dups", s.counts.skipped_duplicates);
  setStatValue("s-errors", s.counts.errors);
  setStatValue("s-removed", s.counts.removed_completed);
  setStatValue("s-updated", s.counts.updated_statuses);

  const isLive = s.state === "running";
  document.getElementById("live-dot-stats").classList.toggle("on", isLive);
  document.getElementById("live-dot-log").classList.toggle("on", isLive);
  document.getElementById("log-box-pretty").classList.toggle("live", isLive);
  document.getElementById("log-box-raw").classList.toggle("live", isLive);

  const btnLaunch = document.getElementById("btn-launch");
  const btnReady = document.getElementById("btn-ready");
  const btnStart = document.getElementById("btn-start");
  const btnStop = document.getElementById("btn-stop");

  btnLaunch.disabled = s.state === "running" || s.state === "launching";
  btnReady.disabled = !(s.state === "waiting_tab" || s.state === "ready");
  btnStart.disabled = s.restart_pending || !(s.state === "ready" || s.state === "stopped" || s.state === "error");
  btnStop.disabled = s.state !== "running" && !s.restart_pending;

  if (!s.selenium_available) {
    btnLaunch.disabled = true;
    btnLaunch.title = "Модуль selenium не установлен";
  }
}

// Небольшая "вспышка" числа при изменении значения - придаёт живости
// статистике на вкладке "Сбор данных", не отвлекая от чтения.
const _lastStatValues = {};
function setStatValue(id, value) {
  const el = document.getElementById(id);
  if (_lastStatValues[id] !== undefined && _lastStatValues[id] !== value) {
    el.classList.remove("flash");
    void el.offsetWidth; // форсируем reflow, чтобы анимация перезапустилась при повторных изменениях
    el.classList.add("flash");
    setTimeout(() => el.classList.remove("flash"), 350);
  }
  _lastStatValues[id] = value;
  el.textContent = value;
}

// Переключатель способа запуска сбора (ссылка / браузер) - оба способа
// равноправно доступны, выбор чисто визуальный (какая панель показана).
document.querySelectorAll(".method-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".method-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".method-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("method-panel-" + btn.dataset.method).classList.add("active");
  });
});

async function pollRunState() {
  try {
    const s = await api("/api/run/state");
    renderRunState(s);
  } catch (e) { /* сервер ещё не готов - пропускаем */ }
}

const LOG_LEVEL_ICONS = {
  success: "✅",
  error: "⚠️",
  warning: "⏸️",
  neutral: "➖",
  info: "•",
};

// Убирает технические префиксы/URL-хвосты и переписывает самые частые
// сообщения человеческим языком - для "простого вида" лога. "Технический"
// вид по-прежнему показывает исходный текст без изменений.
function humanizeLogMessage(msg, level) {
  let text = msg.trim();
  text = text.replace(/^\[IT\/Астана\]\s*/, "");
  text = text.replace(/^\[IT\]\s*/, "");
  text = text.replace(/^\[= дубль\]\s*/, "");
  text = text.replace(/^\[!\]\s*/, "");
  text = text.replace(/^\[авто-перезапуск\]\s*/, "");

  const pageMatch = text.match(/^---\s*Страница\s*(\d+):\s*https?:\/\/\S+/);
  if (pageMatch) return `Проверяю страницу ${pageMatch[1]} результатов поиска…`;

  const idTitleMatch = text.match(/^([^\s:]+):\s*(.+)$/);
  if (level === "success" && idTitleMatch) {
    return `Найден тендер № ${idTitleMatch[1]} — ${idTitleMatch[2]}`;
  }
  if (level === "neutral" && idTitleMatch) {
    return `Уже был в списке: № ${idTitleMatch[1]} — ${idTitleMatch[2]}`;
  }
  return text;
}

function renderPrettyLogEntry(line) {
  if (line.level === "divider") {
    return `<div class="log-entry level-divider"><span class="log-entry-divider-line"></span></div>`;
  }
  const icon = LOG_LEVEL_ICONS[line.level] || "•";
  const text = humanizeLogMessage(line.msg, line.level);
  return `<div class="log-entry level-${line.level}">` +
    `<span class="log-entry-icon">${icon}</span>` +
    `<span class="log-entry-body">${escapeHtml(text)}</span>` +
    `<span class="log-entry-time">${line.ts}</span>` +
  `</div>`;
}

document.querySelectorAll(".log-view-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".log-view-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    const isPretty = btn.dataset.view === "pretty";
    document.getElementById("log-box-pretty").style.display = isPretty ? "block" : "none";
    document.getElementById("log-box-raw").style.display = isPretty ? "none" : "block";
  });
});

async function pollLogs() {
  try {
    const data = await api("/api/run/logs?since=" + lastLogId);
    if (data.lines && data.lines.length) {
      const prettyBox = document.getElementById("log-box-pretty");
      const rawBox = document.getElementById("log-box-raw");
      const prettyAtBottom = prettyBox.scrollTop + prettyBox.clientHeight >= prettyBox.scrollHeight - 20;
      const rawAtBottom = rawBox.scrollTop + rawBox.clientHeight >= rawBox.scrollHeight - 20;

      const emptyPlaceholder = prettyBox.querySelector(".log-empty");
      if (emptyPlaceholder) emptyPlaceholder.remove();

      for (const line of data.lines) {
        prettyBox.insertAdjacentHTML("beforeend", renderPrettyLogEntry(line));
        rawBox.textContent += `[${line.ts}] ${line.msg}\n`;
        lastLogId = line.id;
      }
      if (prettyAtBottom) prettyBox.scrollTop = prettyBox.scrollHeight;
      if (rawAtBottom) rawBox.scrollTop = rawBox.scrollHeight;
    }
  } catch (e) { /* игнорируем временные сбои опроса */ }
}

document.getElementById("auto-restart-checkbox").addEventListener("change", async (e) => {
  await api("/api/run/set_auto_restart", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ enabled: e.target.checked }),
  });
  pollRunState();
});

document.getElementById("btn-save-url").addEventListener("click", async () => {
  const url = document.getElementById("search-url-input").value.trim();
  if (!url) return;
  const res = await api("/api/run/set_url", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ url }),
  });
  if (!res.ok) {
    alert(res.info || "Не удалось сохранить ссылку");
  }
  pollRunState();
});
document.getElementById("search-url-input").addEventListener("keydown", e => {
  if (e.key === "Enter") document.getElementById("btn-save-url").click();
});

document.getElementById("btn-launch").addEventListener("click", async () => {
  await api("/api/run/launch_browser", { method: "POST" });
  pollRunState();
});
document.getElementById("btn-ready").addEventListener("click", async () => {
  await api("/api/run/confirm_ready", { method: "POST" });
  pollRunState();
});
document.querySelectorAll('input[name="collection-filter-mode"]').forEach(radio => {
  radio.addEventListener("change", async () => {
    const res = await api("/api/run/set_filter_mode", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode: radio.value }),
    });
    if (!res.ok) {
      alert(res.info || "Не удалось сменить режим фильтрации");
      pollRunState();
    }
  });
});

document.getElementById("btn-start").addEventListener("click", async () => {
  const selected = document.querySelector('input[name="collection-filter-mode"]:checked');
  if (selected) {
    const modeRes = await api("/api/run/set_filter_mode", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode: selected.value }),
    });
    if (!modeRes.ok) {
      alert(modeRes.info || "Не удалось установить режим фильтрации");
      pollRunState();
      return;
    }
  }
  await api("/api/run/start", { method: "POST" });
  pollRunState();
});
document.getElementById("btn-stop").addEventListener("click", async () => {
  await api("/api/run/stop", { method: "POST" });
  pollRunState();
});

setInterval(pollRunState, 1500);
setInterval(pollLogs, 1500);
pollRunState();
pollLogs();

// ============================================================================
// Вкладка "Тендеры"
// ============================================================================
let tendersPage = 1;
let tendersSortBy = "Начало приёма заявок";
let tendersSortDir = "desc";

function priorityPillClass(p) {
  if (p === "Высокий") return "pill-high";
  if (p === "Средний") return "pill-medium";
  return "pill-low";
}

// Общая логика "срок горит" для Тендеры/Пограничные/Избранные: red — срок
// истёк, yellow — осталось 3 дня или меньше. Единый источник правды и для
// подсветки строки, и для бейджа рядом с датой окончания приёма.
function deadlineRowClass(row) {
  if (row["Просрочен"]) return "row-expired";
  const d = row["ДнейОсталось"];
  if (d !== null && d !== undefined && d <= 3) return "row-warning";
  return "";
}

function deadlineCell(row, endDateValue) {
  const text = escapeHtml(endDateValue || "");
  if (row["Просрочен"]) {
    return `<span class="deadline-pill deadline-pill-expired">Истёк</span>${text ? " " + text : ""}`;
  }
  const d = row["ДнейОсталось"];
  if (d !== null && d !== undefined && d <= 3) {
    const label = d <= 0 ? "сегодня" : (d === 1 ? "1 день" : `${d} дн.`);
    return `<span class="deadline-pill deadline-pill-warning">⏳ ${label}</span>${text ? " " + text : ""}`;
  }
  return text;
}

function fillSelectOptions(selectEl, values, placeholder) {
  const current = selectEl.value;
  selectEl.innerHTML = `<option value="">${placeholder}</option>` +
    values.map(v => `<option value="${v}">${v}</option>`).join("");
  if (values.includes(current)) selectEl.value = current;
}

function updateSortIndicators() {
  document.querySelectorAll("#tenders-table thead th[data-sort]").forEach(th => {
    th.classList.remove("sort-asc", "sort-desc");
    if (th.dataset.sort === tendersSortBy) {
      th.classList.add(tendersSortDir === "asc" ? "sort-asc" : "sort-desc");
    }
  });
}

async function loadTenders() {
  const params = new URLSearchParams({
    search: document.getElementById("f-search").value,
    city: document.getElementById("f-city").value,
    priority: document.getElementById("f-priority").value,
    it_type: document.getElementById("f-ittype").value,
    status: document.getElementById("f-status").value,
    viewed: document.getElementById("f-viewed").value,
    min_amount: document.getElementById("f-min-amount").value,
    sort_by: tendersSortBy,
    sort_dir: tendersSortDir,
    page: tendersPage,
    page_size: 50,
  });
  const data = await api("/api/tenders?" + params.toString());
  updateSortIndicators();

  fillSelectOptions(document.getElementById("f-priority"), data.facets.priority, "Приоритет: все");
  fillSelectOptions(document.getElementById("f-ittype"), data.facets.it_type, "Тип IT: все");
  fillSelectOptions(document.getElementById("f-status"), data.facets.status, "Статус: все");

  const rowsById = {};
  data.rows.forEach(r => { if (r.TenderId) rowsById[r.TenderId] = r; });

  const tbody = document.getElementById("tenders-tbody");
  tbody.innerHTML = data.rows.map(r => `
    <tr class="${r["Просмотрено"] ? "row-viewed" : ""} ${deadlineRowClass(r)}" data-tender-id="${r.TenderId || ""}">
      <td class="viewed-col"><input type="checkbox" class="viewed-checkbox" ${r["Просмотрено"] ? "checked" : ""} ${r.TenderId ? "" : "disabled"}></td>
      <td class="fav-col"><button class="fav-btn ${r["Избранное"] ? "active" : ""}" ${r.TenderId ? "" : "disabled"} title="${r["Избранное"] ? "Убрать из избранного" : "Добавить в избранное"}">${r["Избранное"] ? "★" : "☆"}</button></td>
      <td>${r["№ объявления"] || ""}</td>
      <td class="title-cell">${escapeHtml(r["Название лота/объявления"] || "")}</td>
      <td class="customer-cell">${escapeHtml(r["Заказчик"] || "")}</td>
      <td>${r["Город"] || ""}</td>
      <td>${r["Сумма, тг."] || ""}</td>
      <td>${escapeHtml(r["Статус"] || "")}</td>
      <td>${r["Тип IT"] || ""}</td>
      <td><span class="pill ${priorityPillClass(r["Приоритет"])}">${r["Приоритет"] || ""}</span></td>
      <td>${r["Начало приёма заявок"] || ""}</td>
      <td>${deadlineCell(r, r["Окончание приёма заявок"])}</td>
      <td>${r["Ссылка на объявление"] ? `<a class="tender-link" href="${r["Ссылка на объявление"]}" target="_blank">открыть →</a>` : ""}</td>
      <td>${r.TenderId ? `<button class="btn btn-danger btn-delete-row" type="button">Удалить</button>` : ""}</td>
    </tr>
  `).join("");

  tbody.querySelectorAll(".btn-delete-row").forEach(btn => {
    btn.addEventListener("click", async () => {
      const tr = btn.closest("tr");
      const tenderId = tr.dataset.tenderId;
      if (!confirm("Удалить этот тендер из таблицы? Действие необратимо.")) return;
      btn.disabled = true;
      const res = await api("/api/tenders/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: tenderId }),
      });
      if (res.ok) {
        tr.remove();
      } else {
        btn.disabled = false;
      }
    });
  });

  tbody.querySelectorAll(".fav-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const tr = btn.closest("tr");
      const tenderId = tr.dataset.tenderId;
      if (!tenderId) return;
      const isActive = btn.classList.contains("active");
      btn.disabled = true;
      let res;
      if (isActive) {
        res = await api("/api/favorites/remove", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: tenderId }),
        });
      } else {
        const row = rowsById[tenderId] || {};
        res = await api("/api/favorites", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ ...row, id: tenderId }),
        });
      }
      btn.disabled = false;
      if (res.ok) {
        btn.classList.toggle("active", !isActive);
        btn.textContent = !isActive ? "★" : "☆";
        btn.title = !isActive ? "Убрать из избранного" : "Добавить в избранное";
      }
    });
  });

  tbody.querySelectorAll(".viewed-checkbox").forEach(cb => {
    cb.addEventListener("change", async (e) => {
      const tr = e.target.closest("tr");
      const tenderId = tr.dataset.tenderId;
      if (!tenderId) return;
      const viewed = e.target.checked;
      e.target.disabled = true;
      const res = await api("/api/tenders/viewed", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: tenderId, viewed }),
      });
      e.target.disabled = false;
      if (res.ok) {
        tr.classList.toggle("row-viewed", viewed);
      } else {
        e.target.checked = !viewed; // откатываем чекбокс, если сохранить не удалось
      }
    });
  });

  const totalPages = Math.max(1, Math.ceil(data.total / data.page_size));
  document.getElementById("tenders-page-info").textContent =
    `Страница ${data.page} из ${totalPages} (всего найдено: ${data.total})`;
  document.getElementById("tenders-prev").disabled = data.page <= 1;
  document.getElementById("tenders-next").disabled = data.page >= totalPages;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

document.getElementById("btn-apply-filters").addEventListener("click", () => { tendersPage = 1; loadTenders(); });
document.getElementById("f-viewed").addEventListener("change", () => { tendersPage = 1; loadTenders(); });
document.getElementById("f-search").addEventListener("keydown", e => { if (e.key === "Enter") { tendersPage = 1; loadTenders(); } });
document.getElementById("btn-refresh-tenders").addEventListener("click", async () => {
  await api("/api/tenders/refresh", { method: "POST" });
  loadTenders();
});
document.getElementById("tenders-prev").addEventListener("click", () => { if (tendersPage > 1) { tendersPage--; loadTenders(); } });
document.getElementById("tenders-next").addEventListener("click", () => { tendersPage++; loadTenders(); });

document.querySelectorAll("#tenders-table thead th[data-sort]").forEach(th => {
  th.addEventListener("click", () => {
    const key = th.dataset.sort;
    if (tendersSortBy === key) {
      tendersSortDir = tendersSortDir === "asc" ? "desc" : "asc";
    } else {
      tendersSortBy = key;
      tendersSortDir = "desc";
    }
    loadTenders();
  });
});

// ============================================================================
// Вкладка "Пограничные кандидаты"
// ============================================================================
let borderlinePage = 1;

async function loadBorderline() {
  const params = new URLSearchParams({
    search: document.getElementById("bl-search").value,
    min_score: document.getElementById("bl-min-score").value,
    page: borderlinePage,
    page_size: 50,
  });
  const data = await api("/api/borderline?" + params.toString());

  const tbody = document.getElementById("borderline-tbody");
  tbody.innerHTML = data.rows.map(r => `
    <tr class="${deadlineRowClass(r)}" data-tender-id="${r.tender_id || ""}">
      <td class="fav-col"><button class="fav-btn ${r["Избранное"] ? "active" : ""}" ${r.tender_id ? "" : "disabled"} title="${r["Избранное"] ? "Убрать из избранного" : "Добавить в избранное"}">${r["Избранное"] ? "★" : "☆"}</button></td>
      <td>${r.score}</td>
      <td>${escapeHtml(r.number_anno)}</td>
      <td class="title-cell">${escapeHtml(r.title)}</td>
      <td>${escapeHtml(r.amount || "")}</td>
      <td>${escapeHtml(r.start_date || "")}</td>
      <td>${deadlineCell(r, r.end_date)}</td>
      <td>${escapeHtml(r.keywords)}</td>
      <td>${r.url ? `<a class="tender-link" href="${r.url}" target="_blank">открыть →</a>` : ""}</td>
      <td>${r.ts}</td>
      <td>${r.tender_id ? `<button class="btn btn-danger btn-delete-row" type="button">Удалить</button>` : ""}</td>
    </tr>
  `).join("");

  tbody.querySelectorAll(".btn-delete-row").forEach(btn => {
    btn.addEventListener("click", async () => {
      const tr = btn.closest("tr");
      const tenderId = tr.dataset.tenderId;
      if (!confirm("Удалить этого кандидата из списка? Действие необратимо.")) return;
      btn.disabled = true;
      const res = await api("/api/borderline/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: tenderId }),
      });
      if (res.ok) {
        tr.remove();
      } else {
        btn.disabled = false;
      }
    });
  });

  tbody.querySelectorAll(".fav-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const tr = btn.closest("tr");
      const tenderId = tr.dataset.tenderId;
      if (!tenderId) return;
      const rowData = data.rows.find(r => r.tender_id === tenderId);
      const isActive = btn.classList.contains("active");
      btn.disabled = true;
      let res;
      if (isActive) {
        res = await api("/api/favorites/remove", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: tenderId }),
        });
      } else {
        res = await api("/api/favorites", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            id: tenderId,
            "№ объявления": rowData.number_anno,
            "Название лота/объявления": rowData.title,
            "Сумма, тг.": rowData.amount,
            "Статус": `Пограничный кандидат (балл: ${rowData.score})`,
            "Начало приёма заявок": rowData.start_date,
            "Окончание приёма заявок": rowData.end_date,
            "Ссылка на объявление": rowData.url,
          }),
        });
      }
      btn.disabled = false;
      if (res.ok) {
        btn.classList.toggle("active", !isActive);
        btn.textContent = !isActive ? "★" : "☆";
        btn.title = !isActive ? "Убрать из избранного" : "Добавить в избранное";
      }
    });
  });

  const totalPages = Math.max(1, Math.ceil(data.total / data.page_size));
  document.getElementById("borderline-page-info").textContent =
    `Страница ${data.page} из ${totalPages} (всего: ${data.total})`;
  document.getElementById("borderline-prev").disabled = data.page <= 1;
  document.getElementById("borderline-next").disabled = data.page >= totalPages;
}

document.getElementById("btn-apply-borderline").addEventListener("click", () => { borderlinePage = 1; loadBorderline(); });
document.getElementById("bl-search").addEventListener("keydown", e => { if (e.key === "Enter") { borderlinePage = 1; loadBorderline(); } });
document.getElementById("borderline-prev").addEventListener("click", () => { if (borderlinePage > 1) { borderlinePage--; loadBorderline(); } });
document.getElementById("borderline-next").addEventListener("click", () => { borderlinePage++; loadBorderline(); });

// ============================================================================
// Вкладка "Избранные"
// ============================================================================
let favoriteNoteTimers = {};

async function loadFavorites() {
  const params = new URLSearchParams({ search: document.getElementById("fav-search").value });
  const data = await api("/api/favorites?" + params.toString());

  const tbody = document.getElementById("favorites-tbody");
  const emptyMsg = document.getElementById("favorites-empty");

  if (!data.rows.length) {
    tbody.innerHTML = "";
    emptyMsg.style.display = "block";
    return;
  }
  emptyMsg.style.display = "none";

  tbody.innerHTML = data.rows.map(r => `
    <tr class="${deadlineRowClass(r)}" data-tender-id="${r.TenderId}">
      <td>${escapeHtml(r["№ объявления"] || "")}</td>
      <td class="title-cell">${escapeHtml(r["Название лота/объявления"] || "")}</td>
      <td class="customer-cell">${escapeHtml(r["Заказчик"] || "")}</td>
      <td>${escapeHtml(r["Город"] || "")}</td>
      <td>${escapeHtml(r["Сумма, тг."] || "")}</td>
      <td>${escapeHtml(r["Статус"] || "")}</td>
      <td><span class="pill ${priorityPillClass(r["Приоритет"])}">${r["Приоритет"] || ""}</span></td>
      <td>${escapeHtml(r["Начало приёма заявок"] || "")}</td>
      <td>${deadlineCell(r, r["Окончание приёма заявок"])}</td>
      <td>${r["Ссылка на объявление"] ? `<a class="tender-link" href="${r["Ссылка на объявление"]}" target="_blank">открыть →</a>` : ""}</td>
      <td class="note-col"><input type="text" class="favorite-note-input" placeholder="Заметка..." value="${escapeHtml(r["Заметка"] || "")}"></td>
      <td class="muted">${r["Добавлено"] || ""}</td>
      <td><button class="btn btn-danger btn-remove-favorite" type="button">Убрать</button></td>
    </tr>
  `).join("");

  tbody.querySelectorAll(".favorite-note-input").forEach(input => {
    input.addEventListener("input", (e) => {
      const tenderId = e.target.closest("tr").dataset.tenderId;
      clearTimeout(favoriteNoteTimers[tenderId]);
      favoriteNoteTimers[tenderId] = setTimeout(() => saveFavoriteNote(tenderId, e.target.value), 600);
    });
    input.addEventListener("blur", (e) => {
      const tenderId = e.target.closest("tr").dataset.tenderId;
      clearTimeout(favoriteNoteTimers[tenderId]);
      saveFavoriteNote(tenderId, e.target.value);
    });
  });

  tbody.querySelectorAll(".btn-remove-favorite").forEach(btn => {
    btn.addEventListener("click", async () => {
      const tr = btn.closest("tr");
      const tenderId = tr.dataset.tenderId;
      btn.disabled = true;
      const res = await api("/api/favorites/remove", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: tenderId }),
      });
      if (res.ok) {
        tr.remove();
        if (!document.querySelectorAll("#favorites-tbody tr").length) {
          document.getElementById("favorites-empty").style.display = "block";
        }
      } else {
        btn.disabled = false;
      }
    });
  });
}

async function saveFavoriteNote(tenderId, note) {
  await api("/api/favorites/note", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: tenderId, note }),
  });
}

document.getElementById("btn-apply-favorites").addEventListener("click", loadFavorites);
document.getElementById("fav-search").addEventListener("keydown", e => { if (e.key === "Enter") loadFavorites(); });
document.getElementById("btn-export-favorites").addEventListener("click", () => {
  window.open("/api/favorites/export", "_blank");
});

// ============================================================================
// Вкладка "Ключевые слова" (профили тематик)
// ============================================================================
function listToTextarea(list) { return (list || []).join("\n"); }
function textareaToList(text) {
  return text.split("\n").map(s => s.trim()).filter(Boolean);
}

let profilesState = null;   // { active: [...], profiles: { name: {5 категорий} } }
let editingProfile = null;  // имя профиля, чьё содержимое сейчас в текстовых полях

function setKwStatus(text, ok) {
  const el = document.getElementById("kw-status");
  el.textContent = text;
  el.className = "kw-status" + (ok === true ? " ok" : ok === false ? " err" : "");
}

function setProfileSelectionStatus(text, ok) {
  const el = document.getElementById("profile-selection-status");
  if (!el) return;
  el.textContent = text;
  el.className = "kw-status" + (ok === true ? " ok" : ok === false ? " err" : "");
}

// Что реально сохранено на сервере. Пока пользователь щёлкает чекбоксы,
// меняется только pendingActiveProfiles. Это позволяет выбрать несколько
// тематик и применить их одной кнопкой.
let pendingActiveProfiles = null;

function renderProfileChips() {
  const container = document.getElementById("profile-chips");
  const names = Object.keys(profilesState.profiles);
  if (!Array.isArray(pendingActiveProfiles)) {
    pendingActiveProfiles = Array.from(profilesState.active || []);
  }
  container.innerHTML = names.map(name => `
    <label class="profile-chip ${pendingActiveProfiles.includes(name) ? "active-chip" : ""}" data-profile="${escapeHtml(name)}">
      <input type="checkbox" class="profile-active-checkbox" ${pendingActiveProfiles.includes(name) ? "checked" : ""}>
      ${escapeHtml(name)}
    </label>
  `).join("");

  container.querySelectorAll(".profile-active-checkbox").forEach(cb => {
    cb.addEventListener("change", (e) => {
      const chip = e.target.closest(".profile-chip");
      const name = chip.dataset.profile;
      const set = new Set(pendingActiveProfiles);
      if (e.target.checked) set.add(name); else set.delete(name);
      pendingActiveProfiles = Array.from(set);
      chip.classList.toggle("active-chip", e.target.checked);
      const changed = JSON.stringify(Array.from(profilesState.active || []).sort()) !== JSON.stringify(pendingActiveProfiles.slice().sort());
      setProfileSelectionStatus(changed ? "Есть несохранённые изменения" : "Выбор сохранён", changed ? null : true);
    });
  });
}

async function saveActiveProfiles() {
  const res = await api("/api/keyword-profiles/active", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ active: pendingActiveProfiles }),
  });
  if (res.ok) {
    profilesState = res;
    pendingActiveProfiles = Array.from(res.active || []);
    renderProfileChips();
    setProfileSelectionStatus(
      pendingActiveProfiles.length
        ? `Сохранено: ${pendingActiveProfiles.join(", ")}`
        : "Сохранено: ни одна тематика не выбрана",
      true
    );
  } else {
    setProfileSelectionStatus("Ошибка: " + (res.error || res.info || "не удалось сохранить выбор"), false);
  }
}

function renderProfileSelect() {
  const select = document.getElementById("profile-editor-select");
  const names = Object.keys(profilesState.profiles);
  select.innerHTML = names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
  select.value = editingProfile;
}

function loadEditorFromProfile(name) {
  editingProfile = name;
  const prof = profilesState.profiles[name] || { strong: [], medium: [], weak: [], negative: [], context: [] };
  document.getElementById("kw-strong").value = listToTextarea(prof.strong);
  document.getElementById("kw-medium").value = listToTextarea(prof.medium);
  document.getElementById("kw-weak").value = listToTextarea(prof.weak);
  document.getElementById("kw-context").value = listToTextarea(prof.context);
  document.getElementById("kw-negative").value = listToTextarea(prof.negative);
  document.getElementById("editing-profile-name").textContent = name;
  document.getElementById("profile-editor-select").value = name;
  setKwStatus("", null);
}

async function loadKeywords() {
  profilesState = await api("/api/keyword-profiles");
  pendingActiveProfiles = Array.from(profilesState.active || []);
  setProfileSelectionStatus(
    pendingActiveProfiles.length
      ? `Сохранено: ${pendingActiveProfiles.join(", ")}`
      : "Сохранено: ни одна тематика не выбрана",
    true
  );
  const names = Object.keys(profilesState.profiles);
  if (!editingProfile || !names.includes(editingProfile)) {
    editingProfile = profilesState.active[0] || names[0];
  }
  renderProfileChips();
  renderProfileSelect();
  loadEditorFromProfile(editingProfile);
}

document.getElementById("profile-editor-select").addEventListener("change", (e) => {
  loadEditorFromProfile(e.target.value);
});

document.getElementById("btn-save-active-profiles").addEventListener("click", saveActiveProfiles);

document.getElementById("btn-save-keywords").addEventListener("click", async () => {
  const payload = {
    name: editingProfile,
    keywords: {
      strong: textareaToList(document.getElementById("kw-strong").value),
      medium: textareaToList(document.getElementById("kw-medium").value),
      weak: textareaToList(document.getElementById("kw-weak").value),
      context: textareaToList(document.getElementById("kw-context").value),
      negative: textareaToList(document.getElementById("kw-negative").value),
    },
  };
  const res = await api("/api/keyword-profiles/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (res.ok) {
    profilesState = res;
    setKwStatus("Сохранено. Изменения будут применены к следующему сбору.", true);
  } else {
    setKwStatus("Ошибка: " + (res.error || "не удалось сохранить"), false);
  }
});

document.getElementById("btn-reload-keywords").addEventListener("click", loadKeywords);

document.getElementById("btn-new-profile").addEventListener("click", async () => {
  const name = prompt("Название нового профиля (тематики):");
  if (!name || !name.trim()) return;
  const res = await api("/api/keyword-profiles/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: name.trim(),
      keywords: { strong: [], medium: [], weak: [], negative: [], context: [] },
    }),
  });
  if (res.ok) {
    profilesState = res;
    pendingActiveProfiles = Array.from(res.active || []);
    editingProfile = name.trim();
    renderProfileChips();
    renderProfileSelect();
    loadEditorFromProfile(editingProfile);
    setKwStatus("Профиль создан — не забудьте включить его чипом выше, когда наполните словами.", true);
  }
});

document.getElementById("btn-duplicate-profile").addEventListener("click", async () => {
  const suggested = editingProfile + " (копия)";
  const name = prompt("Название копии профиля:", suggested);
  if (!name || !name.trim()) return;
  const res = await api("/api/keyword-profiles/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: name.trim(),
      keywords: {
        strong: textareaToList(document.getElementById("kw-strong").value),
        medium: textareaToList(document.getElementById("kw-medium").value),
        weak: textareaToList(document.getElementById("kw-weak").value),
        context: textareaToList(document.getElementById("kw-context").value),
        negative: textareaToList(document.getElementById("kw-negative").value),
      },
    }),
  });
  if (res.ok) {
    profilesState = res;
    pendingActiveProfiles = Array.from(res.active || []);
    editingProfile = name.trim();
    renderProfileChips();
    renderProfileSelect();
    loadEditorFromProfile(editingProfile);
  }
});

document.getElementById("btn-rename-profile").addEventListener("click", async () => {
  const newName = prompt("Новое название профиля:", editingProfile);
  if (!newName || !newName.trim() || newName.trim() === editingProfile) return;
  const res = await api("/api/keyword-profiles/rename", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ old_name: editingProfile, new_name: newName.trim() }),
  });
  if (res.ok) {
    profilesState = res;
    pendingActiveProfiles = Array.from(res.active || []);
    editingProfile = newName.trim();
    renderProfileChips();
    renderProfileSelect();
    loadEditorFromProfile(editingProfile);
  }
});

document.getElementById("btn-delete-profile").addEventListener("click", async () => {
  if (Object.keys(profilesState.profiles).length <= 1) {
    alert("Нельзя удалить единственный оставшийся профиль.");
    return;
  }
  if (!confirm(`Удалить профиль «${editingProfile}»? Действие необратимо.`)) return;
  const res = await api("/api/keyword-profiles/delete", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: editingProfile }),
  });
  if (res.ok) {
    profilesState = res;
    pendingActiveProfiles = Array.from(res.active || []);
    editingProfile = Object.keys(profilesState.profiles)[0];
    renderProfileChips();
    renderProfileSelect();
    loadEditorFromProfile(editingProfile);
  }
});

// ============================================================================
// ИИ-проверка релевантности (карточка на вкладке "Ключевые слова")
// ============================================================================
const AI_MODELS_BY_PROVIDER = {
  gemini: [
    { value: "gemini-2.5-flash", label: "Gemini 2.5 Flash (по умолчанию, бесплатно)" },
    { value: "gemini-2.5-flash-lite", label: "Gemini 2.5 Flash-Lite (ещё быстрее/дешевле)" },
  ],
  anthropic: [
    { value: "claude-haiku-4-5-20251001", label: "Claude Haiku 4.5 (быстрее и дешевле)" },
    { value: "claude-sonnet-5", label: "Claude Sonnet 5" },
  ],
};

let aiSettingsState = null;

function populateAiModelOptions(provider, selectedModel) {
  const select = document.getElementById("ai-model");
  const options = AI_MODELS_BY_PROVIDER[provider] || [];
  select.innerHTML = options.map(o => `<option value="${o.value}">${o.label}</option>`).join("");
  if (options.some(o => o.value === selectedModel)) {
    select.value = selectedModel;
  }
}

function updateAiStatusLine() {
  if (!aiSettingsState) return;
  const provider = document.getElementById("ai-provider").value;
  const statusLine = document.getElementById("ai-status-line");
  const hint = document.getElementById("ai-provider-hint");

  if (provider === "gemini") {
    statusLine.textContent = aiSettingsState.gemini_ready
      ? "Ключ найден (GEMINI_API_KEY) — Gemini готов к работе."
      : "Не найден GEMINI_API_KEY в переменных окружения. Получить бесплатный ключ: https://aistudio.google.com/apikey";
    hint.textContent = "Google Gemini даёт постоянный бесплатный тариф без привязки карты (модели Flash/Flash-Lite) — это не то же самое, что лимит сообщений в chatgpt.com/gemini.google.com, тот лимит только для веб-чата человека.";
  } else {
    if (!aiSettingsState.anthropic_installed) {
      statusLine.textContent = "Модуль 'anthropic' не установлен (pip install anthropic).";
    } else if (!aiSettingsState.anthropic_ready) {
      statusLine.textContent = "Не найден ANTHROPIC_API_KEY в переменных окружения.";
    } else {
      statusLine.textContent = "Модуль установлен, ключ найден — Claude готов к работе.";
    }
    hint.textContent = "Anthropic Claude — платный API (как и OpenAI/ChatGPT API): постоянного бесплатного тарифа для API нет ни у одного из них, только небольшой пробный кредит на новый аккаунт.";
  }
}

function updateAiKeyFieldsVisibility(provider) {
  document.getElementById("ai-key-field-gemini").style.display = provider === "gemini" ? "block" : "none";
  document.getElementById("ai-key-field-anthropic").style.display = provider === "anthropic" ? "block" : "none";
}

async function loadAiSettings() {
  aiSettingsState = await api("/api/ai-settings");
  document.getElementById("ai-enabled-checkbox").checked = !!aiSettingsState.enabled;
  document.getElementById("ai-criteria").value = aiSettingsState.criteria || "";
  document.getElementById("ai-verify-mode").value = aiSettingsState.verify_mode || "borderline";
  document.getElementById("ai-provider").value = aiSettingsState.provider || "gemini";
  document.getElementById("ai-gemini-key").value = aiSettingsState.gemini_api_key || "";
  document.getElementById("ai-anthropic-key").value = aiSettingsState.anthropic_api_key || "";
  populateAiModelOptions(aiSettingsState.provider || "gemini", aiSettingsState.model);
  updateAiKeyFieldsVisibility(aiSettingsState.provider || "gemini");
  updateAiStatusLine();
}

document.getElementById("ai-provider").addEventListener("change", (e) => {
  populateAiModelOptions(e.target.value, null);
  updateAiKeyFieldsVisibility(e.target.value);
  updateAiStatusLine();
});

function wireShowHideToggle(buttonId, inputId) {
  document.getElementById(buttonId).addEventListener("click", () => {
    const input = document.getElementById(inputId);
    input.type = input.type === "password" ? "text" : "password";
  });
}
wireShowHideToggle("btn-toggle-gemini-key", "ai-gemini-key");
wireShowHideToggle("btn-toggle-anthropic-key", "ai-anthropic-key");

document.getElementById("btn-save-ai-settings").addEventListener("click", async () => {
  const payload = {
    enabled: document.getElementById("ai-enabled-checkbox").checked,
    criteria: document.getElementById("ai-criteria").value,
    provider: document.getElementById("ai-provider").value,
    verify_mode: document.getElementById("ai-verify-mode").value,
    model: document.getElementById("ai-model").value,
    gemini_api_key: document.getElementById("ai-gemini-key").value,
    anthropic_api_key: document.getElementById("ai-anthropic-key").value,
  };
  const res = await api("/api/ai-settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const statusEl = document.getElementById("ai-save-status");
  if (res.ok) {
    statusEl.textContent = "Сохранено. Применится к следующему сбору.";
    statusEl.className = "kw-status ok";
    aiSettingsState = res;
    updateAiStatusLine();
  } else {
    statusEl.textContent = "Ошибка: " + (res.error || "не удалось сохранить");
    statusEl.className = "kw-status err";
  }
});

// ============================================================================
// Вкладка "Пересмотр ИИ" (разовый прогон по уже собранным данным)
// ============================================================================
let aiRescanPollTimer = null;

function renderAiRescanState(s) {
  const statusLine = document.getElementById("ai-rescan-status-line");
  const progress = document.getElementById("rescan-progress");
  const btn = document.getElementById("btn-start-ai-rescan");

  if (s.running) {
    statusLine.textContent = "Пересмотр идёт...";
    progress.style.display = "block";
    btn.disabled = true;
  } else {
    btn.disabled = false;
    if (s.finished_at) {
      if (s.error) {
        statusLine.textContent = `Завершилось с ошибкой (${s.finished_at}): ${s.error}`;
      } else if (s.checked > 0 && s.failed === s.checked) {
        statusLine.textContent = `Завершено (${s.finished_at}), но ни одна проверка не удалась — см. «Живой лог» на вкладке «Сбор данных» для точной причины (обычно ключ или модель).`;
      } else {
        statusLine.textContent = `Последний пересмотр завершён: ${s.finished_at}`;
      }
      progress.style.display = "block";
    } else {
      statusLine.textContent = "Пересмотр ещё ни разу не запускался.";
      progress.style.display = "none";
    }
  }

  document.getElementById("rescan-checked").textContent = s.checked;
  document.getElementById("rescan-total").textContent = s.total;
  document.getElementById("rescan-promoted").textContent = s.promoted;
  document.getElementById("rescan-flagged").textContent = s.flagged;
  document.getElementById("rescan-failed").textContent = s.failed;
}

async function pollAiRescanState() {
  try {
    const s = await api("/api/ai-rescan/state");
    renderAiRescanState(s);
    clearTimeout(aiRescanPollTimer);
    if (s.running) {
      aiRescanPollTimer = setTimeout(pollAiRescanState, 1500);
    }
  } catch (e) { /* игнорируем временные сбои опроса */ }
}

document.getElementById("btn-start-ai-rescan").addEventListener("click", async () => {
  const targets = [];
  if (document.getElementById("rescan-target-borderline").checked) targets.push("borderline");
  if (document.getElementById("rescan-target-tenders").checked) targets.push("tenders");
  if (!targets.length) {
    alert("Выберите хотя бы одно — «Пограничные» или «Тендеры».");
    return;
  }
  const res = await api("/api/ai-rescan/start", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ targets }),
  });
  if (!res.ok) {
    alert(res.info || "Не удалось запустить пересмотр");
  }
  pollAiRescanState();
});
