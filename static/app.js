const DEFAULT_PREFS = {
  timeFormat: "12h",
  fontSize: "14",
  zoom: "100",
  autocorrect: "on",
  dropdownBehavior: "smart",
  enterKey: "down",
  downKey: "down",
  spaceKey: "right",
  tabKey: "right"
};

let activeInput = null;
let menuHideTimer = null;
let lastSavedValues = new WeakMap();

function loadPrefs() {
  try {
    return { ...DEFAULT_PREFS, ...JSON.parse(localStorage.getItem("gauchoSchedulePrefs") || "{}") };
  } catch {
    return { ...DEFAULT_PREFS };
  }
}

function savePrefs(prefs) {
  localStorage.setItem("gauchoSchedulePrefs", JSON.stringify(prefs));
}

function pad2(value) { return String(value).padStart(2, "0"); }

function formatTime(hour, minute, format) {
  hour = ((hour % 24) + 24) % 24;
  minute = minute || 0;
  if (format === "24h") return `${pad2(hour)}:${pad2(minute)}`;
  return `${hour % 12 || 12}:${pad2(minute)} ${hour < 12 ? "AM" : "PM"}`;
}

function normalizeOneTime(raw, format) {
  let value = String(raw || "").trim();
  if (!value) return "";
  const upper = value.toUpperCase();
  if (upper === "0" || upper === "OFF" || upper === "OOF") return "OFF";
  if (["CL", "CLOSE", "CLOSING"].includes(upper)) return "CLOSE";
  value = value.replace(/[.;]/g, ":");

  let match = value.match(/^\d{3,4}$/);
  if (match) {
    let hour = Number(value.slice(0, -2));
    const minute = Number(value.slice(-2));
    if (hour <= 23 && minute < 60) {
      if (hour <= 7) hour += 12;
      return formatTime(hour, minute, format);
    }
  }

  match = value.match(/^(\d{1,2})(?::(\d{1,2}))?\s*([AP]M)?$/i);
  if (match) {
    let hour = Number(match[1]);
    const minute = Number(match[2] || 0);
    const suffix = match[3] ? match[3].toUpperCase() : null;
    if (Number.isNaN(hour) || Number.isNaN(minute) || minute >= 60) return upper;
    if (suffix === "PM" && hour < 12) hour += 12;
    if (suffix === "AM" && hour === 12) hour = 0;
    if (!suffix && hour <= 7) hour += 12;
    if (hour <= 23) return formatTime(hour, minute, format);
  }
  return upper;
}

function normalizeShift(raw, format) {
  const value = String(raw || "").trim();
  if (!value) return "";
  const upper = value.toUpperCase();
  if (upper === "0" || upper === "OFF" || upper === "OOF") return "OFF";
  if (["CL", "CLOSE", "CLOSING"].includes(upper)) return "CLOSE";
  const pieces = value.split(/\s*(?:-|–|—|;)\s*/).filter(Boolean);
  if (pieces.length === 2) return `${normalizeOneTime(pieces[0], format)}-${normalizeOneTime(pieces[1], format)}`;
  return normalizeOneTime(value, format);
}

function applyVisualPrefs(prefs) {
  document.documentElement.style.setProperty("--editor-font-size", `${prefs.fontSize}px`);
  document.documentElement.style.setProperty("--editor-zoom", String(Number(prefs.zoom) / 100));
  if (document.body.classList.contains("print-page")) {
    document.querySelectorAll(".print-table td").forEach((cell) => {
      const original = cell.dataset.originalValue || cell.textContent.trim();
      cell.dataset.originalValue = original;
      cell.textContent = convertDisplayFormat(original, prefs.timeFormat);
    });
  }
}

function convertDisplayFormat(value, format) {
  if (!value || value === "OFF" || value === "CLOSE") return value;
  return String(value).split("-").map((part) => normalizeOneTime(part, format)).join("-");
}

function focusAndSelect(input) {
  if (!input) return;
  input.focus();
  requestAnimationFrame(() => input.select());
}

function visibleShiftInputs() { return Array.from(document.querySelectorAll(".shift-input")); }

function moveFrom(input, direction) {
  const inputs = visibleShiftInputs();
  const index = inputs.indexOf(input);
  if (index < 0) return;
  const col = Number(input.dataset.col || 0);
  let target = null;
  if (direction === "right") target = inputs[index + 1] || null;
  if (direction === "left") target = inputs[index - 1] || null;
  if (direction === "down") target = inputs.slice(index + 1).find((candidate) => Number(candidate.dataset.col || 0) === col) || null;
  if (direction === "up") target = inputs.slice(0, index).reverse().find((candidate) => Number(candidate.dataset.col || 0) === col) || null;
  focusAndSelect(target);
}

function setAutosaveStatus(text, mode = "saved") {
  const status = document.getElementById("autosave-status");
  if (!status) return;
  status.textContent = text;
  status.dataset.mode = mode;
}

async function postJson(url, payload = {}) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw new Error(await response.text());
  return await response.json();
}

async function patchJson(url, payload = {}) {
  const response = await fetch(url, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw new Error(await response.text());
  return await response.json();
}

function currentWeekStart() {
  return document.querySelector("[data-week-start]")?.dataset.weekStart || null;
}

async function autosaveShift(input, force = false) {
  const weekStart = currentWeekStart();
  if (!weekStart || !input) return;
  const normalized = loadPrefs().autocorrect === "on" ? normalizeShift(input.value, loadPrefs().timeFormat) : input.value.trim();
  input.value = normalized || "OFF";
  if (!force && lastSavedValues.get(input) === input.value) return;

  setAutosaveStatus("Saving…", "saving");
  try {
    const result = await postJson(`/api/week/${weekStart}/shift`, {
      employee_id: input.dataset.employeeId,
      day_index: input.dataset.dayIndex,
      label: input.value
    });
    if (result.label) input.value = result.label;
    lastSavedValues.set(input, input.value);
    setAutosaveStatus("Saved", "saved");
  } catch (error) {
    console.error(error);
    setAutosaveStatus("Autosave failed", "error");
  }
}

function optionsForInput(input) {
  try {
    const parsed = JSON.parse(input.dataset.options || "[]");
    if (Array.isArray(parsed) && parsed.length) return parsed;
  } catch {}
  return ["3:00 PM", "4:00 PM", "5:00 PM", "2:30 PM", "10:00 AM", "10:30 AM", "OFF"];
}

function getSmartMenu() { return document.getElementById("smart-shift-menu"); }

function hideSmartMenu() {
  const menu = getSmartMenu();
  if (menu) menu.hidden = true;
}

function scheduleHideSmartMenu(delay = 220) {
  clearTimeout(menuHideTimer);
  menuHideTimer = setTimeout(hideSmartMenu, delay);
}

function showSmartMenu(input, force = false) {
  clearTimeout(menuHideTimer);
  const prefs = loadPrefs();
  if (prefs.dropdownBehavior === "off") return;
  const menu = getSmartMenu();
  if (!menu) return;
  const basic = ["OFF", "9:00 AM", "10:00 AM", "10:30 AM", "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "2:30 PM", "3:00 PM", "4:00 PM", "5:00 PM", "CLOSE"];
  const rawOptions = prefs.dropdownBehavior === "basic" ? basic : optionsForInput(input);
  const filter = input.value.trim().toUpperCase();
  const options = rawOptions.filter((option) => force || !filter || option.toUpperCase().includes(filter) || filter === "OFF");
  menu.innerHTML = "";
  if (!options.length) { menu.hidden = true; return; }
  options.slice(0, 18).forEach((option, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "smart-option";
    button.innerHTML = `<span>${option}</span>${index === 0 ? '<small>best</small>' : ''}`;
    button.addEventListener("mousedown", async (event) => {
      event.preventDefault();
      clearTimeout(menuHideTimer);
      input.value = option;
      hideSmartMenu();
      focusAndSelect(input);
      await autosaveShift(input, true);
    });
    menu.appendChild(button);
  });
  const rect = input.getBoundingClientRect();
  menu.style.minWidth = `${Math.max(160, rect.width + 26)}px`;
  menu.style.left = `${rect.left + window.scrollX}px`;
  menu.style.top = `${rect.bottom + window.scrollY + 4}px`;
  menu.hidden = false;
}

function setupScheduleEditor() {
  const inputs = visibleShiftInputs();
  if (!inputs.length) return;
  inputs.forEach((input) => {
    lastSavedValues.set(input, input.value);
    input.addEventListener("focus", () => {
      activeInput = input;
      clearTimeout(menuHideTimer);
      input.closest("tr")?.classList.add("active-row");
      input.select();
      showSmartMenu(input, true);
    });
    input.addEventListener("click", () => { input.select(); showSmartMenu(input, true); });
    input.addEventListener("mouseup", (event) => event.preventDefault());
    input.addEventListener("input", () => showSmartMenu(input, false));
    input.addEventListener("blur", async () => {
      input.closest("tr")?.classList.remove("active-row");
      await autosaveShift(input);
      scheduleHideSmartMenu();
    });
    input.addEventListener("keydown", (event) => {
      const prefs = loadPrefs();
      let action = null;
      if (event.key === "Enter") action = prefs.enterKey;
      if (event.key === "ArrowDown") action = prefs.downKey;
      if (event.key === "ArrowUp") action = "up";
      if (event.key === "ArrowLeft") action = "left";
      if (event.key === "ArrowRight") action = "right";
      if (event.key === " " || event.code === "Space") action = prefs.spaceKey;
      if (event.key === "Tab" && prefs.tabKey === "right") action = event.shiftKey ? "left" : "right";
      if (event.key === "Escape") hideSmartMenu();
      if ((event.altKey || event.metaKey) && event.key === "ArrowDown") {
        event.preventDefault();
        showSmartMenu(input, true);
        return;
      }
      if (action && action !== "none" && action !== "browser") {
        event.preventDefault();
        autosaveShift(input, true);
        hideSmartMenu();
        moveFrom(input, action);
      }
    });
  });
  document.querySelectorAll(".cell-menu-button").forEach((button) => {
    button.addEventListener("mousedown", (event) => {
      event.preventDefault();
      clearTimeout(menuHideTimer);
      const input = button.closest(".shift-cell-wrap")?.querySelector(".shift-input");
      if (input) {
        focusAndSelect(input);
        setTimeout(() => showSmartMenu(input, true), 0);
      }
    });
  });
  setupBulkScheduleButtons();
}

function applyShiftPayload(shifts) {
  Object.entries(shifts || {}).forEach(([name, value]) => {
    const input = document.querySelector(`[name="${CSS.escape(name)}"]`);
    if (input) {
      input.value = value;
      lastSavedValues.set(input, value);
    }
  });
}

function setupBulkScheduleButtons() {
  const weekStart = currentWeekStart();
  if (!weekStart) return;
  document.getElementById("fill-best-times")?.addEventListener("click", async () => {
    if (!confirm("Fill every visible schedule cell with its best-known time?")) return;
    setAutosaveStatus("Filling…", "saving");
    try {
      const result = await postJson(`/api/week/${weekStart}/fill-best`, {});
      applyShiftPayload(result.shifts);
      setAutosaveStatus("Saved", "saved");
    } catch (error) {
      console.error(error);
      setAutosaveStatus("Fill failed", "error");
    }
  });
  document.getElementById("copy-previous-week")?.addEventListener("click", async () => {
    if (!confirm("Replace this week with the exact same schedule as the previous week?")) return;
    setAutosaveStatus("Copying…", "saving");
    try {
      const result = await postJson(`/api/week/${weekStart}/copy-previous`, {});
      applyShiftPayload(result.shifts);
      setAutosaveStatus("Saved", "saved");
    } catch (error) {
      console.error(error);
      setAutosaveStatus("Copy failed", "error");
    }
  });
}

function setupPreferencesPage(prefs) {
  const form = document.getElementById("preferences-form");
  if (!form) return;
  form.querySelectorAll("[data-pref]").forEach((field) => { field.value = prefs[field.dataset.pref] ?? DEFAULT_PREFS[field.dataset.pref]; });
  updatePrefOutputs(prefs);
  form.addEventListener("input", () => { const next = collectPrefs(form); updatePrefOutputs(next); applyVisualPrefs(next); });
  document.getElementById("save-preferences")?.addEventListener("click", () => { const next = collectPrefs(form); savePrefs(next); applyVisualPrefs(next); alert("Preferences saved."); });
  document.getElementById("reset-preferences")?.addEventListener("click", () => { savePrefs(DEFAULT_PREFS); window.location.reload(); });
}

function collectPrefs(form) {
  const prefs = { ...DEFAULT_PREFS };
  form.querySelectorAll("[data-pref]").forEach((field) => { prefs[field.dataset.pref] = field.value; });
  return prefs;
}

function updatePrefOutputs(prefs) {
  document.querySelectorAll("[data-pref-output]").forEach((output) => {
    const key = output.dataset.prefOutput;
    output.textContent = `${prefs[key]}${key === "zoom" ? "%" : key === "fontSize" ? "px" : ""}`;
  });
}

function setupSubmitAutocorrect() {
  document.querySelectorAll("form.schedule-form").forEach((form) => {
    form.addEventListener("submit", () => {
      if (loadPrefs().autocorrect !== "on") return;
      form.querySelectorAll(".shift-input").forEach((input) => { input.value = normalizeShift(input.value, loadPrefs().timeFormat); });
    });
  });
}

function debounce(fn, delay = 450) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

function setupEmployeesPage() {
  const board = document.querySelector(".employee-board");
  if (!board) return;
  const saveName = debounce(async (input) => {
    try { await patchJson(`/api/employees/${input.dataset.employeeId}`, { name: input.value }); input.classList.remove("is-saving"); }
    catch (error) { console.error(error); input.classList.add("save-error"); }
  });

  document.querySelectorAll(".employee-name-input").forEach((input) => {
    input.addEventListener("input", () => { input.classList.add("is-saving"); saveName(input); });
    input.addEventListener("blur", () => saveName(input));
  });

  document.getElementById("add-employee-form")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const data = Object.fromEntries(new FormData(form).entries());
    await postJson("/api/employees/add", data);
    window.location.reload();
  });

  document.querySelectorAll(".archive-employee-button").forEach((button) => {
    button.addEventListener("click", async () => {
      if (!confirm("Archive this employee and hide them from the schedule?")) return;
      await postJson(`/api/employees/${button.dataset.employeeId}/archive`, {});
      window.location.reload();
    });
  });
  document.querySelectorAll(".unarchive-employee-button").forEach((button) => {
    button.addEventListener("click", async () => { await postJson(`/api/employees/${button.dataset.employeeId}/unarchive`, {}); window.location.reload(); });
  });

  let dragged = null;
  document.querySelectorAll(".employee-card").forEach((card) => {
    card.addEventListener("dragstart", () => { dragged = card; card.classList.add("dragging"); });
    card.addEventListener("dragend", () => { card.classList.remove("dragging"); dragged = null; saveEmployeeOrder(); });
  });
  document.querySelectorAll(".employee-dropzone").forEach((zone) => {
    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      const after = getDragAfterElement(zone, event.clientY);
      if (!dragged) return;
      if (after == null) zone.appendChild(dragged);
      else zone.insertBefore(dragged, after);
    });
  });
}

function getDragAfterElement(container, y) {
  const cards = [...container.querySelectorAll(".employee-card:not(.dragging)")];
  return cards.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) return { offset, element: child };
    return closest;
  }, { offset: Number.NEGATIVE_INFINITY }).element;
}

async function saveEmployeeOrder() {
  const roles = {};
  document.querySelectorAll(".employee-dropzone").forEach((zone) => {
    roles[zone.dataset.roleId] = [...zone.querySelectorAll(".employee-card")].map((card) => card.dataset.employeeId);
  });
  try { await postJson("/api/employees/reorder", { roles }); }
  catch (error) { console.error(error); alert("Could not save the new employee order."); }
}

document.addEventListener("mousedown", (event) => {
  const menu = getSmartMenu();
  if (!menu) return;
  if (!event.target.closest(".shift-cell-wrap") && !event.target.closest("#smart-shift-menu")) hideSmartMenu();
});

document.addEventListener("DOMContentLoaded", () => {
  const prefs = loadPrefs();
  applyVisualPrefs(prefs);
  setupScheduleEditor();
  setupPreferencesPage(prefs);
  setupSubmitAutocorrect();
  setupEmployeesPage();
});
