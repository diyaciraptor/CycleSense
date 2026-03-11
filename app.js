const STORAGE_KEY = "cyclesense-data-v1";
const THEME_KEY = "cyclesense-theme-v1";
const DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const DEFAULT_DURATION = 5;
const THEMES = ["pink", "blue", "green"];

const state = {
  entries: loadEntries(),
  calendarMonth: startOfMonth(new Date()),
  statusMessage: "No account required. Use XLSX import and export to travel across devices.",
  theme: loadTheme(),
  editingEntryId: null
};

const elements = {
  storageState: document.getElementById("storageState"),
  syncStatus: document.getElementById("syncStatus"),
  resetAppDataButton: document.getElementById("resetAppDataButton"),
  themeSwitcher: document.getElementById("themeSwitcher"),
  form: document.getElementById("cycleForm"),
  editState: document.getElementById("editState"),
  submitButton: document.getElementById("submitButton"),
  startDate: document.getElementById("startDate"),
  endDate: document.getElementById("endDate"),
  flow: document.getElementById("flow"),
  mood: document.getElementById("mood"),
  notes: document.getElementById("notes"),
  nextPrediction: document.getElementById("nextPrediction"),
  predictionNote: document.getElementById("predictionNote"),
  avgCycle: document.getElementById("avgCycle"),
  avgDuration: document.getElementById("avgDuration"),
  cycleCount: document.getElementById("cycleCount"),
  phaseCard: document.getElementById("phaseCard"),
  phaseOrbit: document.getElementById("phaseOrbit"),
  phaseCore: document.getElementById("phaseCore"),
  phaseKicker: document.getElementById("phaseKicker"),
  phaseTitle: document.getElementById("phaseTitle"),
  phaseText: document.getElementById("phaseText"),
  phaseMeta: document.getElementById("phaseMeta"),
  recentCycle: document.getElementById("recentCycle"),
  fertileWindow: document.getElementById("fertileWindow"),
  topSymptoms: document.getElementById("topSymptoms"),
  calendar: document.getElementById("calendar"),
  calendarLabel: document.getElementById("calendarLabel"),
  historyList: document.getElementById("historyList"),
  historyItemTemplate: document.getElementById("historyItemTemplate"),
  logTodayButton: document.getElementById("logTodayButton"),
  resetFormButton: document.getElementById("resetFormButton"),
  prevMonthButton: document.getElementById("prevMonthButton"),
  nextMonthButton: document.getElementById("nextMonthButton"),
  exportButton: document.getElementById("exportButton"),
  importInput: document.getElementById("importInput")
};

bootstrap();

function bootstrap() {
  applyTheme(state.theme);
  hydrateFormDates();
  updateFormMode();
  bindEvents();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./service-worker.js");
  }

  render();
}

function bindEvents() {
  elements.resetAppDataButton.addEventListener("click", handleResetAppData);
  elements.themeSwitcher.addEventListener("click", handleThemeSwitch);
  elements.form.addEventListener("submit", handleSubmit);
  elements.logTodayButton.addEventListener("click", logToday);
  elements.resetFormButton.addEventListener("click", resetForm);
  elements.prevMonthButton.addEventListener("click", () => shiftMonth(-1));
  elements.nextMonthButton.addEventListener("click", () => shiftMonth(1));
  elements.exportButton.addEventListener("click", exportData);
  elements.importInput.addEventListener("change", handleImportXlsx);
}

function loadEntries() {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
    return sanitizeEntries(stored);
  } catch (error) {
    return [];
  }
}

function saveEntries() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.entries));
}

function loadTheme() {
  const storedTheme = localStorage.getItem(THEME_KEY);
  return THEMES.includes(storedTheme) ? storedTheme : "pink";
}

function saveTheme() {
  localStorage.setItem(THEME_KEY, state.theme);
}

function applyTheme(theme) {
  state.theme = THEMES.includes(theme) ? theme : "pink";
  document.body.dataset.theme = state.theme;
  document.querySelectorAll("[data-theme-option]").forEach(button => {
    button.classList.toggle("is-active", button.dataset.themeOption === state.theme);
  });
}

function handleThemeSwitch(event) {
  const button = event.target.closest("[data-theme-option]");
  if (!button) {
    return;
  }

  applyTheme(button.dataset.themeOption);
  saveTheme();
}

function sanitizeEntries(entries) {
  return entries
    .map(normalizeEntry)
    .filter(Boolean)
    .sort((left, right) => parseDateValue(left.startDate) - parseDateValue(right.startDate));
}

function normalizeEntry(entry) {
  if (!entry || !entry.startDate || !entry.endDate) {
    return null;
  }

  const start = parseDateValue(entry.startDate);
  const end = parseDateValue(entry.endDate);
  if (!isValidDate(start) || !isValidDate(end)) {
    return null;
  }

  const normalizedStart = start <= end ? start : end;
  const normalizedEnd = start <= end ? end : start;

  return {
    id: entry.id || createId(),
    startDate: toISODate(normalizedStart),
    endDate: toISODate(normalizedEnd),
    flow: entry.flow || "Medium",
    mood: entry.mood || "Balanced",
    symptoms: Array.isArray(entry.symptoms) ? entry.symptoms : [],
    notes: entry.notes || ""
  };
}

function handleSubmit(event) {
  event.preventDefault();

  const formData = new FormData(elements.form);
  const entry = normalizeEntry({
    id: state.editingEntryId || createId(),
    startDate: formData.get("startDate"),
    endDate: formData.get("endDate"),
    flow: formData.get("flow"),
    mood: formData.get("mood"),
    symptoms: formData.getAll("symptoms"),
    notes: (formData.get("notes") || "").trim()
  });

  if (!entry) {
    alert("Please enter a valid start and end date.");
    return;
  }

  if (state.editingEntryId) {
    state.entries = sanitizeEntries(state.entries.map(item => item.id === state.editingEntryId ? entry : item));
    state.statusMessage = "Cycle updated on this device.";
  } else {
    state.entries = sanitizeEntries(state.entries.concat(entry));
    state.statusMessage = "Cycle saved on this device.";
  }

  saveEntries();
  resetForm();
  render();
}

function logToday() {
  const today = toISODate(new Date());
  elements.startDate.value = today;
  elements.endDate.value = today;
  elements.form.requestSubmit();
}

function resetForm() {
  state.editingEntryId = null;
  elements.form.reset();
  hydrateFormDates();
  updateFormMode();
}

function hydrateFormDates() {
  const today = toISODate(new Date());
  elements.startDate.value = today;
  elements.endDate.value = today;
}

function startEditingEntry(entryId) {
  const entry = state.entries.find(item => item.id === entryId);
  if (!entry) {
    return;
  }

  state.editingEntryId = entry.id;
  elements.startDate.value = entry.startDate;
  elements.endDate.value = entry.endDate;
  elements.flow.value = entry.flow;
  elements.mood.value = entry.mood;
  elements.notes.value = entry.notes;
  elements.form.querySelectorAll('input[name="symptoms"]').forEach(input => {
    input.checked = entry.symptoms.includes(input.value);
  });
  updateFormMode();
  elements.form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function updateFormMode() {
  const isEditing = Boolean(state.editingEntryId);

  if (elements.editState) {
    elements.editState.hidden = !isEditing;
  }
  if (elements.submitButton) {
    elements.submitButton.textContent = isEditing ? "Update cycle" : "Save cycle";
  }
  if (elements.resetFormButton) {
    elements.resetFormButton.textContent = isEditing ? "Cancel edit" : "Reset";
  }
}

function shiftMonth(offset) {
  const next = new Date(state.calendarMonth);
  next.setMonth(next.getMonth() + offset);
  state.calendarMonth = startOfMonth(next);
  renderCalendar();
}

function handleResetAppData() {
  const shouldReset = window.confirm("Clear all saved cycles from this device?");
  if (!shouldReset) {
    return;
  }

  state.entries = [];
  saveEntries();
  resetForm();
  state.statusMessage = "App data cleared on this device.";
  render();
}

function render() {
  renderTransferState();
  renderStats();
  renderPhase();
  renderInsights();
  renderCalendar();
  renderHistory();
}

function renderTransferState() {
  elements.storageState.textContent = state.entries.length
    ? `${state.entries.length} cycle${state.entries.length === 1 ? "" : "s"} stored on this device`
    : "Stored on this device";
  elements.syncStatus.textContent = state.statusMessage;
}

function renderStats() {
  const metrics = getMetrics();
  const nextStart = metrics.predictedStart;

  elements.nextPrediction.textContent = nextStart
    ? formatLongDate(nextStart)
    : "Add two cycles to unlock predictions";
  elements.predictionNote.textContent = nextStart
    ? `${daysFromToday(nextStart)} from today based on your average cycle length`
    : "All data stays on this device unless you export it.";
  elements.avgCycle.textContent = metrics.averageCycle ? `${metrics.averageCycle} days` : "--";
  elements.avgDuration.textContent = metrics.averageDuration ? `${metrics.averageDuration} days` : "--";
  elements.cycleCount.textContent = String(state.entries.length);
}

function renderInsights() {
  const metrics = getMetrics();
  const latest = state.entries[state.entries.length - 1];

  elements.recentCycle.textContent = latest
    ? `${formatShortDate(latest.startDate)} to ${formatShortDate(latest.endDate)} - ${diffInDays(latest.startDate, latest.endDate) + 1} days`
    : "No data yet";

  elements.fertileWindow.textContent = metrics.fertileWindow
    ? `${formatShortDate(metrics.fertileWindow.start)} to ${formatShortDate(metrics.fertileWindow.end)}`
    : "Needs at least two logs";

  elements.topSymptoms.textContent = metrics.topSymptoms.length
    ? metrics.topSymptoms.join(", ")
    : "Start logging to see patterns";
}

function renderPhase() {
  const phase = getCurrentPhase(getMetrics());

  elements.phaseCard.dataset.phase = phase.id;
  elements.phaseKicker.textContent = phase.kicker;
  elements.phaseTitle.textContent = phase.title;
  elements.phaseText.textContent = phase.description;
  if (elements.phaseMeta) {
    elements.phaseMeta.textContent = phase.meta || "Prediction updates as you log more cycles.";
  }
  if (elements.phaseOrbit) {
    elements.phaseOrbit.style.setProperty("--phase-progress", `${phase.progress || 0}%`);
  }
  elements.phaseCore.setAttribute("aria-label", phase.title);
}

function renderCalendar() {
  const metrics = getMetrics();
  const monthStart = startOfMonth(state.calendarMonth);
  const firstGridDate = startOfWeek(monthStart);
  const today = toISODate(new Date());
  const periodDates = new Set(expandPeriodDates(state.entries));
  const predictedDates = new Set(expandPredictedDates(metrics.predictedStart, metrics.averageDuration));
  const fertileDates = new Set(expandDateRange(metrics.fertileWindow));

  elements.calendar.innerHTML = "";
  elements.calendarLabel.textContent = monthStart.toLocaleDateString(undefined, {
    month: "long",
    year: "numeric"
  });

  DAY_NAMES.forEach(day => {
    const heading = document.createElement("div");
    heading.className = "weekday";
    heading.textContent = day;
    elements.calendar.appendChild(heading);
  });

  for (let index = 0; index < 42; index += 1) {
    const date = addDays(firstGridDate, index);
    const isoDate = toISODate(date);
    const cell = document.createElement("div");
    const tags = document.createElement("div");
    const number = document.createElement("div");

    cell.className = "day";
    if (date.getMonth() !== monthStart.getMonth()) {
      cell.classList.add("day-muted");
    }
    if (isoDate === today) {
      cell.classList.add("day-today");
    }

    number.className = "day-number";
    number.textContent = String(date.getDate());
    cell.appendChild(number);

    tags.className = "day-tags";

    const isPeriodDay = periodDates.has(isoDate);
    const isPredictedDay = predictedDates.has(isoDate);
    const isFertileDay = fertileDates.has(isoDate);

    if (isPeriodDay) {
      tags.appendChild(createTag("Period", "tag-period"));
    }
    if (!isPeriodDay && isPredictedDay) {
      tags.appendChild(createTag("Predicted", "tag-predicted"));
    }
    if (!isPeriodDay && !isPredictedDay && isFertileDay) {
      tags.appendChild(createTag("Fertile", "tag-fertile"));
    }

    cell.appendChild(tags);
    elements.calendar.appendChild(cell);
  }
}

function renderHistory() {
  elements.historyList.innerHTML = "";

  if (!state.entries.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No cycles saved yet. Add your first cycle or import an XLSX backup.";
    elements.historyList.appendChild(empty);
    return;
  }

  state.entries
    .slice()
    .reverse()
    .forEach(entry => {
      const fragment = elements.historyItemTemplate.content.cloneNode(true);
      const container = fragment.querySelector(".history-item");

      fragment.querySelector(".history-dates").textContent =
        `${formatLongDate(entry.startDate)} to ${formatLongDate(entry.endDate)}`;
      fragment.querySelector(".history-meta").textContent =
        `${diffInDays(entry.startDate, entry.endDate) + 1} days - ${entry.flow} flow - ${entry.mood}`;
      fragment.querySelector(".history-notes").textContent =
        entry.symptoms.length || entry.notes
          ? [entry.symptoms.join(", "), entry.notes].filter(Boolean).join(" - ")
          : "No extra notes";

      container.querySelector(".edit-button").addEventListener("click", () => {
        startEditingEntry(entry.id);
      });

      container.querySelector(".delete-button").addEventListener("click", () => {
        state.entries = state.entries.filter(item => item.id !== entry.id);
        if (state.editingEntryId === entry.id) {
          resetForm();
        }
        saveEntries();
        state.statusMessage = "Cycle deleted from this device.";
        render();
      });

      elements.historyList.appendChild(fragment);
    });
}

function getMetrics() {
  const cycleLengths = [];
  const symptomCounter = new Map();

  state.entries.forEach((entry, index) => {
    entry.symptoms.forEach(symptom => {
      symptomCounter.set(symptom, (symptomCounter.get(symptom) || 0) + 1);
    });

    if (index > 0) {
      cycleLengths.push(diffInDays(state.entries[index - 1].startDate, entry.startDate));
    }
  });

  const averageCycle = cycleLengths.length ? Math.round(average(cycleLengths)) : null;
  const durations = state.entries.map(entry => diffInDays(entry.startDate, entry.endDate) + 1);
  const averageDuration = durations.length ? Math.round(average(durations)) : null;
  const latestEntry = state.entries[state.entries.length - 1] || null;
  const latestStart = latestEntry ? parseDateValue(latestEntry.startDate) : null;
  const predictedStart = averageCycle && latestStart
    ? addDays(latestStart, averageCycle)
    : null;
  const predictedOvulation = predictedStart
    ? addDays(predictedStart, -14)
    : null;
  const fertileWindow = predictedOvulation
    ? {
        start: addDays(predictedOvulation, -5),
        end: predictedOvulation
      }
    : null;

  const topSymptoms = [...symptomCounter.entries()]
    .sort((left, right) => right[1] - left[1])
    .slice(0, 3)
    .map(([symptom]) => symptom);

  return {
    averageCycle,
    averageDuration,
    predictedStart,
    predictedOvulation,
    fertileWindow,
    topSymptoms
  };
}

function getCurrentPhase(metrics) {
  const today = parseDateValue(new Date());
  const latest = state.entries[state.entries.length - 1];
  const estimatedCycleLength = clamp(metrics.averageCycle || 28, 21, 35);
  const estimatedPeriodLength = clamp(metrics.averageDuration || DEFAULT_DURATION, 3, 7);

  if (!latest) {
    return {
      id: "unknown",
      kicker: "Cycle estimate",
      title: "Not enough cycle data yet",
      description: "Log your period dates and CycleSense will estimate your current menstrual, follicular, ovulation, and luteal phases.",
      meta: "Predictions get better as you log more full cycles.",
      progress: 18
    };
  }

  const latestStart = parseDateValue(latest.startDate);
  const latestEnd = parseDateValue(latest.endDate);
  const cycleDay = Math.max(1, diffInDays(latestStart, today) + 1);
  const predictedNextStart = metrics.predictedStart || addDays(latestStart, estimatedCycleLength);
  const predictedOvulation = metrics.predictedOvulation || addDays(predictedNextStart, -14);
  const ovulationWindowStart = addDays(predictedOvulation, -1);
  const ovulationWindowEnd = addDays(predictedOvulation, 1);
  const follicularStart = latestStart;
  const follicularEnd = addDays(predictedOvulation, -2);
  const lutealStart = addDays(predictedOvulation, 1);

  if (today >= latestStart && today <= latestEnd) {
    const loggedPeriodLength = Math.max(diffInDays(latestStart, latestEnd) + 1, 1);
    const elapsedDays = diffInDays(latestStart, today) + 1;
    const progress = clamp(Math.round((elapsedDays / loggedPeriodLength) * 100), 8, 100);

    return {
      id: "period",
      kicker: "Menstrual phase",
      title: "You appear to be in the menstrual phase",
      description: "Cleveland Clinic describes menses as the first phase of the cycle, beginning on day 1 of bleeding. The follicular phase also starts on day 1, so these two phases overlap here.",
      meta: `Cycle day ${cycleDay} - logged bleed ${loggedPeriodLength} day${loggedPeriodLength === 1 ? "" : "s"}`,
      progress
    };
  }

  if (today < ovulationWindowStart) {
    const follicularSpan = Math.max(diffInDays(follicularStart, follicularEnd), 1);
    const progress = clamp(Math.round((diffInDays(follicularStart, today) / follicularSpan) * 100), 8, 100);

    return {
      id: "follicular",
      kicker: "Follicular phase",
      title: "You appear to be in the follicular phase",
      description: "Cleveland Clinic describes the follicular phase as starting on the first day of your period and lasting until ovulation. After bleeding ends, this card stays in the follicular phase until your estimated ovulation window.",
      meta: `Cycle day ${cycleDay} - ovulation estimated around ${formatShortDate(predictedOvulation)}`,
      progress
    };
  }

  if (today >= ovulationWindowStart && today <= ovulationWindowEnd) {
    const ovulationSpan = Math.max(diffInDays(ovulationWindowStart, ovulationWindowEnd), 1);
    const progress = clamp(Math.round((diffInDays(ovulationWindowStart, today) / ovulationSpan) * 100), 8, 100);

    return {
      id: "ovulation",
      kicker: "Ovulation phase",
      title: "You appear to be near ovulation",
      description: "Cleveland Clinic describes ovulation as happening roughly around day 14 in a 28-day cycle. CycleSense treats this as a short estimated window centered on your predicted ovulation date.",
      meta: `Estimated ovulation ${formatShortDate(predictedOvulation)} - fertile window ${formatShortDate(addDays(predictedOvulation, -5))} to ${formatShortDate(predictedOvulation)}`,
      progress
    };
  }

  if (today < predictedNextStart) {
    const lutealSpan = Math.max(diffInDays(lutealStart, predictedNextStart), 1);
    const progress = clamp(Math.round((diffInDays(lutealStart, today) / lutealSpan) * 100), 8, 100);
    const daysUntilNext = diffInDays(today, predictedNextStart);

    return {
      id: "luteal",
      kicker: "Luteal phase",
      title: "You appear to be in the luteal phase",
      description: "Cleveland Clinic describes the luteal phase as the time from about day 15 until the next period. This is the phase after ovulation and before your next predicted start date.",
      meta: `${daysUntilNext} day${daysUntilNext === 1 ? "" : "s"} until predicted next period - next start ${formatShortDate(predictedNextStart)}`,
      progress
    };
  }

  return {
    id: "luteal",
    kicker: "Late luteal estimate",
    title: "Your next period may be due",
    description: "Your current date is at or beyond the predicted next start, so CycleSense is holding the end of the luteal estimate until you log your next period.",
    meta: `Predicted next start ${formatShortDate(predictedNextStart)} - log your next period to refresh the estimate`,
    progress: 100
  };
}

function exportData() {
  const blob = buildWorkbookBlob(state.entries);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `cyclesense-export-${toISODate(new Date())}.xlsx`;
  link.click();
  URL.revokeObjectURL(url);
  state.statusMessage = "XLSX exported. Import this file on another device to move your data.";
  renderTransferState();
}

async function handleImportXlsx(event) {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  try {
    const entries = await parseWorkbookFile(file);
    state.entries = sanitizeEntries(entries);
    saveEntries();
    resetForm();
    state.statusMessage = `Imported ${state.entries.length} cycle${state.entries.length === 1 ? "" : "s"} from XLSX.`;
    render();
  } catch (error) {
    alert(`That XLSX file could not be imported: ${error.message}`);
  }

  event.target.value = "";
}

async function parseWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const files = await unzipEntries(buffer);
  const worksheetXml = files.get("xl/worksheets/sheet1.xml");

  if (!worksheetXml) {
    throw new Error("CycleSense sheet not found.");
  }

  const sharedStringsXml = files.get("xl/sharedStrings.xml") || "";
  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const rows = parseWorksheetRows(worksheetXml, sharedStrings);

  if (rows.length <= 1) {
    return [];
  }

  return rows.slice(1).map(row => ({
    id: row[7] || createId(),
    startDate: normalizeImportedDate(row[0]),
    endDate: normalizeImportedDate(row[1]),
    flow: row[3] || "Medium",
    mood: row[4] || "Balanced",
    symptoms: row[5] ? row[5].split(",").map(value => value.trim()).filter(Boolean) : [],
    notes: row[6] || ""
  }));
}

async function unzipEntries(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const entries = new Map();
  let offset = 0;

  while (offset + 4 <= bytes.length) {
    const signature = readUInt32(bytes, offset);
    if (signature !== 0x04034b50) {
      break;
    }

    const compressionMethod = readUInt16(bytes, offset + 8);
    const compressedSize = readUInt32(bytes, offset + 18);
    const fileNameLength = readUInt16(bytes, offset + 26);
    const extraLength = readUInt16(bytes, offset + 28);
    const fileNameStart = offset + 30;
    const fileNameEnd = fileNameStart + fileNameLength;
    const dataStart = fileNameEnd + extraLength;
    const dataEnd = dataStart + compressedSize;
    const fileName = decodeUtf8(bytes.slice(fileNameStart, fileNameEnd));
    const fileData = bytes.slice(dataStart, dataEnd);

    if (compressionMethod === 0) {
      entries.set(fileName, decodeUtf8(fileData));
    } else if (compressionMethod === 8) {
      const decompressed = await inflateRaw(fileData);
      entries.set(fileName, decodeUtf8(new Uint8Array(decompressed)));
    }

    offset = dataEnd;
  }

  return entries;
}

function readUInt16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUInt32(bytes, offset) {
  return readUInt16(bytes, offset) | (readUInt16(bytes, offset + 2) << 16);
}

function decodeUtf8(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return await new Response(stream).arrayBuffer();
}

function parseSharedStrings(xmlText) {
  if (!xmlText) {
    return [];
  }

  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map(item =>
    Array.from(item.getElementsByTagName("t")).map(node => node.textContent || "").join("")
  );
}

function parseWorksheetRows(xmlText, sharedStrings) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  const rows = [];

  Array.from(doc.getElementsByTagName("row")).forEach(rowNode => {
    const row = [];
    Array.from(rowNode.getElementsByTagName("c")).forEach(cell => {
      const reference = cell.getAttribute("r") || "A1";
      const index = columnIndex(reference.replace(/\d+/g, ""));
      const type = cell.getAttribute("t") || "";
      let value = "";

      if (type === "inlineStr") {
        value = Array.from(cell.getElementsByTagName("t")).map(node => node.textContent || "").join("");
      } else {
        const rawValue = cell.getElementsByTagName("v")[0]?.textContent || "";
        value = type === "s" ? (sharedStrings[Number(rawValue)] || "") : rawValue;
      }

      row[index] = value;
    });
    rows.push(row.map(value => value || ""));
  });

  return rows;
}

function columnIndex(name) {
  let index = 0;
  for (let position = 0; position < name.length; position += 1) {
    index = index * 26 + (name.charCodeAt(position) - 64);
  }
  return index - 1;
}

function normalizeImportedDate(value) {
  if (!value) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  if (/^\d+$/.test(String(value).trim())) {
    const date = addDays(new Date(1899, 11, 30), Number(value));
    return toISODate(date);
  }

  return toISODate(parseDateValue(value));
}

function buildWorkbookBlob(entries) {
  const files = [
    {
      name: "[Content_Types].xml",
      content: createContentTypesXml()
    },
    {
      name: "_rels/.rels",
      content: createRootRelsXml()
    },
    {
      name: "xl/workbook.xml",
      content: createWorkbookXml()
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: createWorkbookRelsXml()
    },
    {
      name: "xl/styles.xml",
      content: createStylesXml()
    },
    {
      name: "xl/worksheets/sheet1.xml",
      content: createWorksheetXml(entries)
    }
  ];

  const zipBytes = buildZipArchive(files);
  return new Blob([zipBytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function createWorksheetXml(entries) {
  const headers = [
    "Start Date",
    "End Date",
    "Duration (days)",
    "Flow",
    "Mood",
    "Symptoms",
    "Notes",
    "Entry ID"
  ];

  const rows = [headers].concat(entries.map(entry => [
    entry.startDate,
    entry.endDate,
    String(diffInDays(entry.startDate, entry.endDate) + 1),
    entry.flow,
    entry.mood,
    entry.symptoms.join(", "),
    entry.notes,
    entry.id
  ]));

  const sheetRows = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndexValue) => {
      const cellReference = `${columnName(columnIndexValue)}${rowIndex + 1}`;
      return `<c r="${cellReference}" t="inlineStr"><is><t>${escapeXml(value || "")}</t></is></c>`;
    }).join("");

    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");

  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>",
    "<sheetFormatPr defaultRowHeight=\"15\"/>",
    `<dimension ref=\"A1:H${rows.length}\"/>`,
    "<cols>",
    "<col min=\"1\" max=\"2\" width=\"16\" customWidth=\"1\"/>",
    "<col min=\"3\" max=\"3\" width=\"14\" customWidth=\"1\"/>",
    "<col min=\"4\" max=\"5\" width=\"18\" customWidth=\"1\"/>",
    "<col min=\"6\" max=\"6\" width=\"24\" customWidth=\"1\"/>",
    "<col min=\"7\" max=\"7\" width=\"36\" customWidth=\"1\"/>",
    "<col min=\"8\" max=\"8\" width=\"24\" customWidth=\"1\"/>",
    "</cols>",
    `<sheetData>${sheetRows}</sheetData>`,
    "</worksheet>"
  ].join("");
}

function createContentTypesXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
    "</Types>"
  ].join("");
}

function createRootRelsXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>",
    "</Relationships>"
  ].join("");
}

function createWorkbookXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
    "<sheets><sheet name=\"Cycles\" sheetId=\"1\" r:id=\"rId1\"/></sheets>",
    "</workbook>"
  ].join("");
}

function createWorkbookRelsXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
    "</Relationships>"
  ].join("");
}

function createStylesXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>",
    "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
    "</styleSheet>"
  ].join("");
}

function buildZipArchive(files) {
  const encoder = new TextEncoder();
  const fileRecords = files.map(file => {
    const nameBytes = encoder.encode(file.name);
    const contentBytes = encoder.encode(file.content);
    return {
      nameBytes,
      contentBytes,
      crc32: crc32(contentBytes)
    };
  });

  let offset = 0;
  const localChunks = [];
  const centralChunks = [];

  fileRecords.forEach(file => {
    const localHeader = createLocalFileHeader(file.nameBytes, file.contentBytes.length, file.crc32);
    localChunks.push(localHeader, file.nameBytes, file.contentBytes);

    const centralHeader = createCentralDirectoryHeader(file.nameBytes, file.contentBytes.length, file.crc32, offset);
    centralChunks.push(centralHeader, file.nameBytes);

    offset += localHeader.length + file.nameBytes.length + file.contentBytes.length;
  });

  const centralDirectory = concatUint8Arrays(centralChunks);
  const endRecord = createEndOfCentralDirectoryRecord(fileRecords.length, centralDirectory.length, offset);

  return concatUint8Arrays(localChunks.concat([centralDirectory, endRecord]));
}

function createLocalFileHeader(nameBytes, size, checksum) {
  const header = new Uint8Array(30);
  const view = new DataView(header.buffer);

  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint32(14, checksum >>> 0, true);
  view.setUint32(18, size, true);
  view.setUint32(22, size, true);
  view.setUint16(26, nameBytes.length, true);
  view.setUint16(28, 0, true);

  return header;
}

function createCentralDirectoryHeader(nameBytes, size, checksum, offset) {
  const header = new Uint8Array(46);
  const view = new DataView(header.buffer);

  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 0, true);
  view.setUint32(16, checksum >>> 0, true);
  view.setUint32(20, size, true);
  view.setUint32(24, size, true);
  view.setUint16(28, nameBytes.length, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, offset, true);

  return header;
}

function createEndOfCentralDirectoryRecord(fileCount, centralDirectorySize, centralDirectoryOffset) {
  const record = new Uint8Array(22);
  const view = new DataView(record.buffer);

  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, fileCount, true);
  view.setUint16(10, fileCount, true);
  view.setUint32(12, centralDirectorySize, true);
  view.setUint32(16, centralDirectoryOffset, true);
  view.setUint16(20, 0, true);

  return record;
}

function crc32(bytes) {
  let crc = 0xffffffff;

  bytes.forEach(byte => {
    crc ^= byte;
    for (let index = 0; index < 8; index += 1) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  });

  return (crc ^ 0xffffffff) >>> 0;
}

function concatUint8Arrays(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;

  chunks.forEach(chunk => {
    result.set(chunk, offset);
    offset += chunk.length;
  });

  return result;
}

function columnName(index) {
  let name = "";
  let value = index;

  while (value >= 0) {
    name = String.fromCharCode((value % 26) + 65) + name;
    value = Math.floor(value / 26) - 1;
  }

  return name;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function expandPeriodDates(entries) {
  return entries.flatMap(entry => expandDateRange({
    start: parseDateValue(entry.startDate),
    end: parseDateValue(entry.endDate)
  }));
}

function expandPredictedDates(predictedStart, averageDuration) {
  if (!predictedStart) {
    return [];
  }

  const duration = Math.max((averageDuration || DEFAULT_DURATION) - 1, 0);
  return expandDateRange({
    start: predictedStart,
    end: addDays(predictedStart, duration)
  });
}

function expandDateRange(range) {
  if (!range || !range.start || !range.end) {
    return [];
  }

  const start = parseDateValue(range.start);
  const end = parseDateValue(range.end);
  if (!isValidDate(start) || !isValidDate(end) || start > end) {
    return [];
  }

  const dayCount = diffInDays(start, end);
  return Array.from({ length: dayCount + 1 }, (_, index) => toISODate(addDays(start, index)));
}

function createTag(text, className) {
  const tag = document.createElement("span");
  tag.className = `tag ${className}`;
  tag.textContent = text;
  return tag;
}

function toISODate(date) {
  const copy = parseDateValue(date);
  const year = copy.getFullYear();
  const month = String(copy.getMonth() + 1).padStart(2, "0");
  const day = String(copy.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateValue(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "string") {
    const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    }
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return new Date(NaN);
  }

  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function isValidDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function startOfWeek(date) {
  return addDays(date, -date.getDay());
}

function addDays(date, amount) {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  return next;
}

function diffInDays(startDate, endDate) {
  const start = parseDateValue(startDate);
  const end = parseDateValue(endDate);
  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  return Math.round((end - start) / millisecondsPerDay);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function average(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function formatLongDate(date) {
  return parseDateValue(date).toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric"
  });
}

function formatShortDate(date) {
  return parseDateValue(date).toLocaleDateString(undefined, {
    month: "short",
    day: "numeric"
  });
}

function daysFromToday(date) {
  const today = toISODate(new Date());
  const diff = diffInDays(today, toISODate(date));

  if (diff === 0) {
    return "today";
  }
  if (diff === 1) {
    return "tomorrow";
  }
  if (diff > 1) {
    return `in ${diff} days`;
  }

  return `${Math.abs(diff)} days ago`;
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }

  return `entry-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}








