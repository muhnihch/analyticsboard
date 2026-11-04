const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8UJmjPQty11jTPW88byAPgm1ZKHgf2qZsKQWPG-opwJ1kW55AzwNnjvGO9EsyxRkG/exec";

const CHART_LIMIT = 12;
const COLORS = ["#1f6f5b", "#d08a2d", "#3d6f8f", "#8c4a3a", "#6b705c", "#7a6a9a"];
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const elements = {
  updated: document.getElementById("updated"),
  status: document.getElementById("status"),
  recalculate: document.getElementById("recalculate"),
  fileDrop: document.getElementById("file-drop"),
  fileTitle: document.getElementById("file-title"),
  fileInput: document.getElementById("data-file"),
  fileName: document.getElementById("file-name"),
  uploadResult: document.getElementById("upload-result"),
  preview: document.getElementById("upload-preview"),
  previewSummary: document.getElementById("preview-summary"),
  previewColumns: document.getElementById("preview-columns"),
  previewActions: document.getElementById("preview-actions"),
  previewAdd: document.getElementById("preview-add"),
  previewCancel: document.getElementById("preview-cancel"),
  uploadWait: document.getElementById("upload-wait"),
  uploadWaitText: document.getElementById("upload-wait-text"),
  deleteData: document.getElementById("delete-data"),
  deleteConfirm: document.getElementById("delete-confirm"),
  deleteCancel: document.getElementById("delete-cancel"),
  deleteCopy: document.getElementById("delete-copy"),
  dataCount: document.getElementById("data-count"),
  dataTable: document.getElementById("data-table"),
  descriptive: document.getElementById("descriptive"),
  correlation: document.getElementById("correlation"),
  frequencies: document.getElementById("frequencies"),
  figures: document.getElementById("figures"),
  figuresEmpty: document.getElementById("figures-empty"),
  barSelect: document.getElementById("bar-column"),
  pieSelect: document.getElementById("pie-column"),
  barTitle: document.getElementById("bar-title"),
  pieTitle: document.getElementById("pie-title"),
  barCanvas: document.getElementById("bar-chart"),
  pieCanvas: document.getElementById("pie-chart")
};

const numberFormat = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const countFormat = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const percentFormat = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });

let summary = null;
let pendingUpload = null;
let busy = false;
const charts = { bar: null, pie: null };

if (window.Chart) {
  Chart.defaults.font.family = '"Source Sans 3", "Segoe UI", Helvetica, Arial, sans-serif';
  Chart.defaults.color = "#5c6670";
}

elements.recalculate.addEventListener("click", function () {
  loadSummary(true);
});
elements.fileInput.addEventListener("change", previewUpload);
elements.previewAdd.addEventListener("click", sendUpload);
elements.previewCancel.addEventListener("click", function () {
  clearUpload();
  setStatus("Upload canceled. The current data is unchanged.");
});
elements.deleteData.addEventListener("click", function () {
  elements.deleteData.hidden = true;
  elements.deleteConfirm.hidden = false;
  elements.deleteCancel.hidden = false;
  elements.deleteCopy.textContent = "This removes every row and column from the sheet. You cannot undo this.";
});
elements.deleteCancel.addEventListener("click", resetDelete);
elements.deleteConfirm.addEventListener("click", deleteStoredData);
elements.barSelect.addEventListener("change", drawCharts);
elements.pieSelect.addEventListener("change", drawCharts);

init();

function init() {
  if (!isConfigured()) {
    elements.updated.textContent = "Add your web app URL to app.js to load the summary.";
    setStatus("You can choose a file and preview it. Add your Apps Script /exec URL in app.js before replacing the sheet.", true);
    renderEmptySections();
    setBusy(false);
    return;
  }
  loadSummary(false);
}

function isConfigured() {
  return /^https:\/\/script\.google\.com\/.+\/exec$/.test(APPS_SCRIPT_URL);
}

async function loadSummary(rebuild) {
  setBusy(true);
  setStatus(rebuild ? "Recalculating from the Data tab." : "Loading the latest summary.");
  try {
    const data = await fetchSummary(rebuild);
    render(data);
    setStatus(rebuild ? "Summary recalculated." : "");
  } catch (error) {
    console.error(error);
    setStatus("The summary could not be loaded. Check your connection and the web app URL, then try again.", true);
    if (!summary) {
      elements.updated.textContent = "No summary loaded yet.";
      renderEmptySections();
    }
  } finally {
    setBusy(false);
  }
}

async function fetchSummary(rebuild) {
  const url = APPS_SCRIPT_URL + (rebuild ? "?rebuild=1" : "");
  const response = await fetch(url, { cache: "no-store" });
  const data = await response.json();
  if (!response.ok || !data || data.error) {
    throw new Error((data && data.error) || "Request failed with status " + response.status);
  }
  if (!Array.isArray(data.columns) || !Array.isArray(data.correlations)) {
    throw new Error("The response is not a summary.");
  }
  return data;
}

async function postRecords(payload) {
  const response = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload)
  });
  const result = await response.json();
  if (!response.ok || !result || result.error) {
    throw new Error((result && result.error) || "Request failed with status " + response.status);
  }
  return result;
}

function render(data) {
  summary = normalizeSummary(data);
  elements.updated.textContent = "Updated " + formatDateTime(summary.updatedAt) + ". " +
    countFormat.format(summary.rowCount) + (summary.rowCount === 1 ? " row" : " rows") + " in the sheet.";
  renderDataTable(columnNames(), summary.rows);
  elements.descriptive.innerHTML = renderDescriptive();
  elements.correlation.innerHTML = renderCorrelations();
  elements.frequencies.innerHTML = renderFrequencies();
  fillChartSelects();
  drawCharts();
}

function normalizeSummary(data) {
  const columns = data.columns.map(function (column) {
    const copy = Object.assign({}, column);
    if (copy.type === "word") {
      if (!Number.isFinite(Number(copy.distinct)) || copy.distinct === null) {
        copy.distinct = Array.isArray(copy.counts) ? copy.counts.length : 0;
      } else {
        copy.distinct = Number(copy.distinct);
      }
    }
    return copy;
  });
  return {
    updatedAt: data.updatedAt,
    rowCount: Number(data.rowCount) || 0,
    columns: columns,
    correlations: data.correlations,
    rows: Array.isArray(data.rows) ? data.rows : []
  };
}

function renderEmptySections() {
  elements.descriptive.innerHTML = '<p class="message">Descriptive statistics appear here once the summary loads.</p>';
  elements.correlation.innerHTML = '<p class="message">Correlations appear here once the summary loads.</p>';
  elements.frequencies.innerHTML = '<p class="message">Frequencies appear here once the summary loads.</p>';
  renderDataTable([], []);
  elements.figures.hidden = true;
  elements.figuresEmpty.hidden = false;
  elements.figuresEmpty.textContent = "Charts appear here once the summary loads.";
}

function numberColumns() {
  return summary.columns.filter(function (column) { return column.type === "number"; });
}

function wordColumns() {
  return summary.columns.filter(function (column) { return column.type === "word"; });
}

function chartableColumns() {
  return wordColumns().filter(function (column) {
    return column.n > 0 && column.distinct <= CHART_LIMIT && Array.isArray(column.counts) && column.counts.length > 0;
  });
}

function columnNames() {
  if (!summary) return [];
  if (summary.columns.length) return summary.columns.map(function (column) { return column.name; });
  const names = [];
  (summary.rows || []).forEach(function (row) {
    Object.keys(row).forEach(function (key) {
      if (names.indexOf(key) === -1) names.push(key);
    });
  });
  return names;
}

function renderDataTable(columns, rows) {
  elements.dataTable.innerHTML = "";
  const count = rows.length;
  elements.dataCount.textContent = count
    ? countFormat.format(count) + (count === 1 ? " row" : " rows")
    : "Import a spreadsheet to see the rows.";
  if (!count) {
    elements.dataTable.innerHTML = '<p class="message">No rows to show.</p>';
    return;
  }

  const scroll = document.createElement("div");
  scroll.className = "data-scroll";
  const table = document.createElement("table");
  table.className = "data-table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  const indexHead = document.createElement("th");
  indexHead.className = "row-index";
  indexHead.textContent = "#";
  headRow.appendChild(indexHead);
  columns.forEach(function (name) {
    const cell = document.createElement("th");
    cell.textContent = name;
    headRow.appendChild(cell);
  });
  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  rows.forEach(function (row, index) {
    const tableRow = document.createElement("tr");
    const indexCell = document.createElement("td");
    indexCell.className = "row-index";
    indexCell.textContent = String(index + 1);
    tableRow.appendChild(indexCell);
    columns.forEach(function (name) {
      const cell = document.createElement("td");
      const value = row[name];
      if (typeof value === "number" && Number.isFinite(value)) cell.className = "numeric";
      cell.textContent = value == null ? "" : String(value);
      tableRow.appendChild(cell);
    });
    tbody.appendChild(tableRow);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  scroll.appendChild(table);
  elements.dataTable.appendChild(scroll);
}

function resetDelete() {
  elements.deleteData.hidden = false;
  elements.deleteConfirm.hidden = true;
  elements.deleteCancel.hidden = true;
  elements.deleteCopy.textContent = "Remove every row and column currently stored in the sheet.";
}

async function deleteStoredData() {
  if (busy || !isConfigured()) return;
  setBusy(true);
  setStatus("Deleting the stored data.");
  try {
    await postRecords({ action: "delete" });
    clearUpload();
    resetDelete();
    await refreshAfterSave("The stored data was deleted.");
  } catch (error) {
    console.error(error);
    setStatus("The data could not be deleted. Click Recalculate to see what is still stored.", true);
  } finally {
    setBusy(false);
  }
}

async function previewUpload() {
  const file = elements.fileInput.files[0];
  pendingUpload = null;
  elements.preview.hidden = true;
  resetUploadSuccess();
  if (!file) {
    elements.fileName.textContent = "No file selected";
    return;
  }
  elements.fileTitle.textContent = "File selected";
  elements.fileName.textContent = file.name;

  if (!/\.(xlsx|csv)$/i.test(file.name)) {
    setStatus("Choose an .xlsx or .csv file.", true);
    clearUpload();
    return;
  }
  if (!window.XLSX) {
    setStatus("The spreadsheet reader did not load. Check your connection and reload the page.", true);
    clearUpload();
    return;
  }

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true }) : [];
    const prepared = prepareUpload(rows);

    if (!prepared.records.length) {
      setStatus("That file has no data rows to add.", true);
      clearUpload();
      return;
    }

    pendingUpload = prepared;
    showPreview(prepared);
    setStatus("");
  } catch (error) {
    console.error(error);
    setStatus("The file could not be read. Use a valid .xlsx or .csv file.", true);
    clearUpload();
  }
}

function prepareUpload(rows) {
  const columns = [];
  const seen = {};
  const records = [];
  let skipped = 0;

  rows.forEach(function (row) {
    const record = {};
    let hasValue = false;

    Object.keys(row).forEach(function (rawKey) {
      const key = String(rawKey).trim();
      if (!key || /^__EMPTY/.test(key)) return;
      const lower = key.toLowerCase();
      const name = key;

      if (!Object.prototype.hasOwnProperty.call(seen, lower)) {
        seen[lower] = true;
        columns.push(name);
      }

      const value = cleanUploadValue(row[rawKey]);
      if (value === "") return;
      if (Object.prototype.hasOwnProperty.call(record, name)) return;
      record[name] = value;
      hasValue = true;
    });

    if (hasValue) records.push(record);
    else skipped += 1;
  });

  return { records: records, columns: columns, skipped: skipped };
}

function cleanUploadValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : formatDate(value, true);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  return String(value).trim();
}

function showPreview(prepared) {
  const count = prepared.records.length;
  let text = "This replaces the current data. " + countFormat.format(count) +
    (count === 1 ? " row will be imported." : " rows will be imported.");
  if (prepared.skipped) {
    text += " " + countFormat.format(prepared.skipped) +
      (prepared.skipped === 1 ? " empty row will be skipped." : " empty rows will be skipped.");
  }
  elements.previewSummary.textContent = text;
  elements.previewColumns.textContent = prepared.columns.join(", ");
  elements.preview.hidden = false;
  renderDataTable(prepared.columns, prepared.records);
  elements.dataCount.textContent = countFormat.format(prepared.records.length) +
    (prepared.records.length === 1 ? " row in this file." : " rows in this file.") +
    " Choose Replace data to store them.";
}

async function sendUpload() {
  if (busy || !pendingUpload) return;
  const records = pendingUpload.records;
  const fileName = elements.fileName.textContent;
  setBusy(true);
  showUploadWait(fileName);
  setStatus("Uploading " + fileName + ".");
  try {
    const result = await postRecords({ replace: true, records: records });
    const inserted = Number(result.inserted) || 0;
    const message = uploadMessage(fileName, inserted);
    await refreshAfterSave(message);
    clearUpload();
    showUploadSuccess(fileName, message);
  } catch (error) {
    console.error(error);
    hideUploadWait();
    setStatus("The sheet could not be replaced. Click Recalculate to see what is currently saved, then try again.", true);
  } finally {
    hideUploadWait();
    setBusy(false);
  }
}

async function refreshAfterSave(message) {
  try {
    render(await fetchSummary(false));
    setStatus(message);
  } catch (error) {
    console.error(error);
    setStatus(message + " The summary could not be refreshed. Click Recalculate to update it.", true);
  }
}

function uploadMessage(fileName, inserted) {
  const count = countFormat.format(inserted);
  const stored = inserted === 1 ? "1 row is now stored." : count + " rows are now stored.";
  return fileName + " was uploaded. " + stored;
}

function showUploadWait(fileName) {
  elements.preview.hidden = false;
  elements.preview.classList.add("is-waiting");
  elements.previewActions.hidden = true;
  elements.uploadWait.hidden = false;
  elements.uploadWaitText.textContent = "Uploading " + fileName + ".";
}

function hideUploadWait() {
  elements.preview.classList.remove("is-waiting");
  elements.previewActions.hidden = false;
  elements.uploadWait.hidden = true;
}

function showUploadSuccess(fileName, message) {
  elements.fileDrop.classList.add("is-uploaded");
  elements.fileTitle.textContent = "Uploaded";
  elements.fileName.textContent = fileName;
  elements.uploadResult.hidden = false;
  elements.uploadResult.textContent = message;
}

function resetUploadSuccess() {
  elements.fileDrop.classList.remove("is-uploaded");
  elements.fileTitle.textContent = "Choose a file";
  elements.uploadResult.hidden = true;
  elements.uploadResult.textContent = "";
}

function clearUpload() {
  pendingUpload = null;
  elements.fileInput.value = "";
  hideUploadWait();
  resetUploadSuccess();
  elements.fileName.textContent = "No file selected";
  elements.preview.hidden = true;
  renderDataTable(columnNames(), summary && summary.rows ? summary.rows : []);
}

function renderDescriptive() {
  const numbers = numberColumns();
  if (!numbers.length) {
    return '<p class="message">Add a number column to see descriptive statistics.</p>';
  }

  const body = numbers.map(function (column) {
    return "<tr><td class=\"stub\">" + escapeHtml(column.name) + "</td><td>" + countFormat.format(column.n) +
      "</td><td>" + formatStat(column.mean) + "</td><td>" + formatStat(column.median) +
      "</td><td>" + formatStat(column.sd) + "</td></tr>";
  }).join("");

  return apaTable({
    number: "Table 1",
    title: "Descriptive Statistics for Numeric Variables",
    header: "<tr><th class=\"stub\">Variable</th><th><em>N</em></th><th><em>M</em></th><th><em>Mdn</em></th><th><em>SD</em></th></tr>",
    body: body,
    note: "<em>Note.</em> <em>N</em> = number of nonblank values. <em>M</em> = mean; <em>Mdn</em> = median; <em>SD</em> = sample standard deviation."
  });
}

function renderCorrelations() {
  const numbers = numberColumns();
  if (numbers.length < 2) {
    return '<p class="message">Add at least two number columns to see correlations.</p>';
  }

  const header = "<tr><th class=\"stub\">Variable</th>" + numbers.map(function (_, index) {
    return "<th>" + (index + 1) + "</th>";
  }).join("") + "</tr>";

  const body = numbers.map(function (rowColumn, rowIndex) {
    const cells = numbers.map(function (column, columnIndex) {
      if (rowIndex === columnIndex) return "<td>—</td>";
      if (columnIndex > rowIndex) return "<td></td>";
      const match = findCorrelation(rowColumn.name, column.name);
      return "<td>" + formatR(match ? match.r : null) + "</td>";
    }).join("");
    return "<tr><td class=\"stub\">" + (rowIndex + 1) + ". " + escapeHtml(rowColumn.name) + "</td>" + cells + "</tr>";
  }).join("");

  const counts = summary.correlations
    .map(function (item) { return Number(item.n); })
    .filter(Number.isFinite);
  let nText = "";
  if (counts.length) {
    const smallest = Math.min.apply(null, counts);
    const largest = Math.max.apply(null, counts);
    nText = smallest === largest
      ? "<em>N</em> = " + countFormat.format(smallest) + " for each pair. "
      : "Pairwise <em>N</em> ranges from " + countFormat.format(smallest) + " to " + countFormat.format(largest) + ". ";
  }

  return apaTable({
    number: "Table 2",
    title: "Pearson Correlations Among Numeric Variables",
    header: header,
    body: body,
    note: "<em>Note.</em> " + nText + "Each pair uses rows where both values are filled. A dash means fewer than 3 pairs or no variation."
  });
}

function renderFrequencies() {
  const words = wordColumns().filter(function (column) { return column.n > 0; });
  if (!words.length) {
    return '<p class="message">Add a text column to see frequencies.</p>';
  }

  const tabulated = chartableColumns();
  const skipped = words.filter(function (column) { return tabulated.indexOf(column) === -1; });

  let html = "";
  if (tabulated.length) {
    const body = tabulated.map(function (column) {
      return column.counts.map(function (item, index) {
        const percent = column.n ? (item.count / column.n) * 100 : 0;
        return "<tr" + (index === 0 ? " class=\"group-start\"" : "") + "><td class=\"stub\">" +
          (index === 0 ? escapeHtml(column.name) : "") + "</td><td class=\"stub\">" + escapeHtml(item.value) +
          "</td><td>" + countFormat.format(item.count) + "</td><td>" + percentFormat.format(percent) + "%</td></tr>";
      }).join("");
    }).join("");

    html += apaTable({
      number: "Table 3",
      title: "Frequencies of Categorical Variables",
      header: "<tr><th class=\"stub\">Variable</th><th class=\"stub\">Response</th><th><em>n</em></th><th>%</th></tr>",
      body: body,
      note: "<em>Note.</em> Percentages use the nonblank answers in each column and may not total 100% because of rounding."
    });
  }

  if (skipped.length) {
    const list = skipped.map(function (column) {
      return escapeHtml(column.name) + " (" + countFormat.format(column.distinct) + " different answers)";
    }).join(", ");
    html += "<p class=\"note\">Not tabulated because they have more than " + CHART_LIMIT + " different answers: " + list + ".</p>";
  }

  return html;
}

function apaTable(parts) {
  return "<table class=\"apa\"><caption><span class=\"table-number\">" + parts.number +
    "</span><span class=\"table-title\">" + parts.title + "</span></caption><thead>" +
    parts.header + "</thead><tbody>" + parts.body + "</tbody></table><p class=\"note\">" + parts.note + "</p>";
}

function fillChartSelects() {
  const eligible = chartableColumns();
  const previousBar = elements.barSelect.value;
  const previousPie = elements.pieSelect.value;

  [elements.barSelect, elements.pieSelect].forEach(function (select) {
    select.innerHTML = "";
    eligible.forEach(function (column) {
      const option = document.createElement("option");
      option.value = column.name;
      option.textContent = column.name;
      select.appendChild(option);
    });
  });

  elements.figures.hidden = eligible.length === 0;
  elements.figuresEmpty.hidden = eligible.length > 0;
  elements.figuresEmpty.textContent = "Add a text column with " + CHART_LIMIT + " or fewer different answers to see charts.";
  if (!eligible.length) return;

  const varied = eligible.filter(function (column) { return column.counts.length > 1; });
  const defaultBar = varied[0] || eligible[0];
  const defaultPie = varied[1] || eligible.find(function (column) { return column !== defaultBar; }) || defaultBar;

  elements.barSelect.value = hasColumn(eligible, previousBar) ? previousBar : defaultBar.name;
  elements.pieSelect.value = hasColumn(eligible, previousPie) ? previousPie : defaultPie.name;
}

function drawCharts() {
  drawChart("bar", elements.barSelect.value, elements.barCanvas, elements.barTitle);
  drawChart("pie", elements.pieSelect.value, elements.pieCanvas, elements.pieTitle);
}

function drawChart(kind, columnName, canvas, titleNode) {
  if (charts[kind]) {
    charts[kind].destroy();
    charts[kind] = null;
  }
  if (!summary || !window.Chart) return;

  const column = chartableColumns().find(function (item) { return item.name === columnName; });
  if (!column) {
    titleNode.textContent = kind === "bar" ? "Bar chart" : "Pie chart";
    return;
  }

  titleNode.textContent = column.name;
  const labels = column.counts.map(function (item) { return item.value; });
  const values = column.counts.map(function (item) { return item.count; });
  const colors = labels.map(function (_, index) { return COLORS[index % COLORS.length]; });

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: kind === "pie", position: "bottom", labels: { boxWidth: 12, padding: 14 } },
      tooltip: {
        callbacks: {
          label: function (context) {
            const count = context.parsed && typeof context.parsed === "object" ? context.parsed.y : context.parsed;
            const percent = column.n ? (count / column.n) * 100 : 0;
            return " " + countFormat.format(count) + " (" + percentFormat.format(percent) + "%)";
          }
        }
      }
    }
  };

  if (kind === "bar") {
    options.scales = {
      x: { grid: { display: false } },
      y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#efe9df" } }
    };
  }

  charts[kind] = new Chart(canvas, {
    type: kind,
    data: {
      labels: labels,
      datasets: [{
        label: "Count",
        data: values,
        backgroundColor: colors,
        borderColor: kind === "pie" ? "#ffffff" : colors,
        borderWidth: kind === "pie" ? 2 : 0,
        borderRadius: kind === "bar" ? 6 : 0,
        maxBarThickness: 64
      }]
    },
    options: options
  });
}

function findCorrelation(a, b) {
  return summary.correlations.find(function (item) {
    return (item.variableA === a && item.variableB === b) || (item.variableA === b && item.variableB === a);
  });
}

function hasColumn(columns, name) {
  return columns.some(function (column) { return column.name === name; });
}

function formatStat(value) {
  if (value === null || value === undefined || value === "") return "—";
  const number = Number(value);
  return Number.isFinite(number) ? numberFormat.format(number) : "—";
}

function formatR(value) {
  if (value === null || value === undefined || value === "") return "—";
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  const text = Math.abs(number).toFixed(2).replace(/^0/, "");
  return number < 0 && text !== ".00" ? "-" + text : text;
}

function formatDate(date, includeTime) {
  let text = MONTHS[date.getMonth()] + " " + date.getDate() + ", " + date.getFullYear();
  if (includeTime && (date.getHours() || date.getMinutes())) text += " at " + formatTime(date);
  return text;
}

function formatTime(date) {
  const hours = date.getHours() % 12 || 12;
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return hours + ":" + minutes + " " + (date.getHours() >= 12 ? "PM" : "AM");
}

function formatDateTime(iso) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "recently";
  return formatDate(date, false) + " at " + formatTime(date);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function setStatus(message, isError) {
  elements.status.textContent = message;
  elements.status.hidden = !message;
  elements.status.classList.toggle("is-error", Boolean(isError));
}

function setBusy(value) {
  busy = value;
  const locked = value || !isConfigured();
  [elements.recalculate, elements.previewAdd, elements.deleteData, elements.deleteConfirm].forEach(function (button) {
    button.disabled = locked;
  });
  elements.previewCancel.disabled = value;
  elements.fileInput.disabled = value;
}
