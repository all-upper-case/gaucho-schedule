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

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatTime(hour, minute, format) {
  hour = ((hour % 24) + 24) % 24;
  minute = minute || 0;
  if (format === "24h") {
    return `${pad2(hour)}:${pad2(minute)}`;
  }
  const suffix = hour < 12 ? "AM" : "PM";
  let h12 = hour % 12;
  if (h12 === 0) h12 = 12;
  return `${h12}:${pad2(minute)} ${suffix}`;
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
    let minute = Number(value.slice(-2));
    if (hour <= 23 && minute < 60) {
      if (hour <= 7) hour += 12;
      return formatTime(hour, minute, format);
    }
  }

  match = value.match(/^(\d{1,2})(?::(\d{1,2}))?\s*([AP]M)?$/i);
  if (match) {
    let hour = Number(match[1]);
    let minute = Number(match[2] || 0);
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
  if (pieces.length === 2) {
    return `${normalizeOneTime(pieces[0], format)}-${normalizeOneTime(pieces[1], format)}`;
  }
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

function visibleShiftInputs() {
  return Array.from(document.querySelectorAll(".shift-input"));
}

function moveFrom(input, direction) {
  const inputs = visibleShiftInputs();
  const index = inputs.indexOf(input);
  if (index < 0) return;
  const col = Number(input.dataset.col || 0);
  let target = null;

  if (direction === "right") {
    target = inputs[index + 1] || null;
  } else if (direction === "left") {
    target = inputs[index - 1] || null;
  } else if (direction === "down") {
    target = inputs.slice(index + 1).find((candidate) => Number(candidate.dataset.col || 0) === col) || null;
  } else if (direction === "up") {
    target = inputs.slice(0, index).reverse().find((candidate) => Number(candidate.dataset.col || 0) === col) || null;
  }

  focusAndSelect(target);
}

function setupScheduleEditor(prefs) {
  const inputs = visibleShiftInputs();
  if (!inputs.length) return;

  inputs.forEach((input) => {
    input.addEventListener("focus", () => input.select());
    input.addEventListener("click", () => input.select());
    input.addEventListener("mouseup", (event) => event.preventDefault());

    input.addEventListener("blur", () => {
      if (loadPrefs().autocorrect === "on") {
        input.value = normalizeShift(input.value, loadPrefs().timeFormat);
      }
    });

    input.addEventListener("keydown", (event) => {
      const currentPrefs = loadPrefs();
      let action = null;

      if (event.key === "Enter") action = currentPrefs.enterKey;
      if (event.key === "ArrowDown") action = currentPrefs.downKey;
      if (event.key === "ArrowUp") action = "up";
      if (event.key === "ArrowLeft") action = "left";
      if (event.key === "ArrowRight") action = "right";
      if (event.key === " " || event.code === "Space") action = currentPrefs.spaceKey;
      if (event.key === "Tab" && currentPrefs.tabKey === "right") action = event.shiftKey ? "left" : "right";

      if (action && action !== "none" && action !== "browser") {
        event.preventDefault();
        if (currentPrefs.autocorrect === "on") {
          input.value = normalizeShift(input.value, currentPrefs.timeFormat);
        }
        moveFrom(input, action);
      }
    });
  });

  applyDropdownPreference(prefs);
}

function applyDropdownPreference(prefs) {
  const inputs = visibleShiftInputs();
  inputs.forEach((input) => {
    if (prefs.dropdownBehavior === "off") {
      input.dataset.originalList = input.dataset.originalList || input.getAttribute("list") || "";
      input.removeAttribute("list");
    } else {
      const original = input.dataset.originalList;
      if (original) input.setAttribute("list", original);
    }
  });
}

function setupPreferencesPage(prefs) {
  const form = document.getElementById("preferences-form");
  if (!form) return;

  form.querySelectorAll("[data-pref]").forEach((field) => {
    field.value = prefs[field.dataset.pref] ?? DEFAULT_PREFS[field.dataset.pref];
  });
  updatePrefOutputs(prefs);

  form.addEventListener("input", () => {
    const next = collectPrefs(form);
    updatePrefOutputs(next);
    applyVisualPrefs(next);
  });

  document.getElementById("save-preferences")?.addEventListener("click", () => {
    const next = collectPrefs(form);
    savePrefs(next);
    applyVisualPrefs(next);
    alert("Preferences saved.");
  });

  document.getElementById("reset-preferences")?.addEventListener("click", () => {
    savePrefs(DEFAULT_PREFS);
    window.location.reload();
  });
}

function collectPrefs(form) {
  const prefs = { ...DEFAULT_PREFS };
  form.querySelectorAll("[data-pref]").forEach((field) => {
    prefs[field.dataset.pref] = field.value;
  });
  return prefs;
}

function updatePrefOutputs(prefs) {
  document.querySelectorAll("[data-pref-output]").forEach((output) => {
    const key = output.dataset.prefOutput;
    const suffix = key === "zoom" ? "%" : key === "fontSize" ? "px" : "";
    output.textContent = `${prefs[key]}${suffix}`;
  });
}

function setupSubmitAutocorrect() {
  document.querySelectorAll("form.schedule-form").forEach((form) => {
    form.addEventListener("submit", () => {
      const prefs = loadPrefs();
      if (prefs.autocorrect !== "on") return;
      form.querySelectorAll(".shift-input").forEach((input) => {
        input.value = normalizeShift(input.value, prefs.timeFormat);
      });
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  const prefs = loadPrefs();
  applyVisualPrefs(prefs);
  setupScheduleEditor(prefs);
  setupPreferencesPage(prefs);
  setupSubmitAutocorrect();
});
