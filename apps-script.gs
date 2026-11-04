var DATA_SHEET = 'Data';
var RESULTS_SHEET = 'Results';
var MAX_STORED_COUNTS = 50;
var LOCK_WAIT_MS = 30000;
var RESULTS_WIDTH = 7;

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var summary = params.rebuild === '1' ? rebuildSummary() : readSummary_();
    if (!summary) summary = rebuildSummary();
    if (!summary.rows) summary.rows = readDataRows_();
    return json_(summary);
  } catch (error) {
    return json_({ error: String(error.message || error) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('The request had no data.');
    var payload = JSON.parse(e.postData.contents);

    lock.waitLock(LOCK_WAIT_MS);
    if (payload && payload.action === 'delete') {
      getSheet_(DATA_SHEET).clear();
      return json_({ ok: true, deleted: true, summary: buildSummary_() });
    }

    var replacing = payload && payload.replace === true && !Array.isArray(payload);
    var records = replacing ? payload.records : (Array.isArray(payload) ? payload : [payload]);
    if (!Array.isArray(records)) throw new Error('The request had no rows.');

    var inserted = replacing ? replaceRecords_(records) : appendRecords_(records);
    var summary = buildSummary_();
    return json_({ ok: true, inserted: inserted, replaced: replacing, summary: summary });
  } catch (error) {
    return json_({ error: String(error.message || error) });
  } finally {
    lock.releaseLock();
  }
}

function rebuildSummary() {
  var lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT_MS);
  try {
    return buildSummary_();
  } finally {
    lock.releaseLock();
  }
}

function buildSummary_() {
  var sheet = getSheet_(DATA_SHEET);
  var values = sheet.getLastRow() === 0 ? [] : sheet.getDataRange().getValues();
  var headers = values.length ? trimHeaders_(values[0]) : [];

  var dataRows = [];
  for (var r = 1; r < values.length; r++) {
    var hasValue = false;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] && !isBlank_(values[r][i])) hasValue = true;
    }
    if (hasValue) dataRows.push(values[r]);
  }

  var columns = [];
  var numberColumns = [];

  for (var col = 0; col < headers.length; col++) {
    if (!headers[col]) continue;
    var cells = dataRows.map(function (row) { return row[col]; });
    var first = firstFilled_(cells);
    var isNumber = first !== null && toNumber_(first) !== null;

    if (isNumber) {
      var numbers = cells.map(toNumber_);
      var present = numbers.filter(function (value) { return value !== null; });
      columns.push({
        name: headers[col],
        type: 'number',
        n: present.length,
        mean: present.length ? round6_(mean_(present)) : null,
        median: present.length ? round6_(median_(present)) : null,
        sd: present.length > 1 ? round6_(sampleSd_(present)) : null
      });
      numberColumns.push({ name: headers[col], values: numbers });
    } else {
      var counts = {};
      var order = [];
      var answered = 0;
      cells.forEach(function (cell) {
        var word = toWord_(cell);
        if (word === null) return;
        answered++;
        if (!has_(counts, word)) {
          counts[word] = 0;
          order.push(word);
        }
        counts[word]++;
      });
      var entry = {
        name: headers[col],
        type: 'word',
        n: answered,
        mean: null,
        median: null,
        sd: null,
        distinct: order.length
      };
      if (order.length <= MAX_STORED_COUNTS) {
        entry.counts = order.map(function (value) {
          return { value: value, count: counts[value] };
        }).sort(function (a, b) { return b.count - a.count; });
      }
      columns.push(entry);
    }
  }

  var correlations = [];
  for (var a = 0; a < numberColumns.length; a++) {
    for (var b = a + 1; b < numberColumns.length; b++) {
      var result = pearson_(numberColumns[a].values, numberColumns[b].values);
      correlations.push({
        variableA: numberColumns[a].name,
        variableB: numberColumns[b].name,
        r: result.r === null ? null : round6_(result.r),
        n: result.n
      });
    }
  }

  var summary = {
    updatedAt: new Date().toISOString(),
    rowCount: dataRows.length,
    columns: columns,
    correlations: correlations,
    rows: rowsFrom_(headers, dataRows)
  };
  writeResults_(summary);
  return summary;
}

function readDataRows_() {
  var sheet = getSheet_(DATA_SHEET);
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getDisplayValues();
  var headers = trimHeaders_(values[0]);
  var dataRows = [];
  for (var r = 1; r < values.length; r++) dataRows.push(values[r]);
  return rowsFrom_(headers, dataRows).filter(function (row) {
    return headers.some(function (header) { return header && String(row[header]).trim() !== ''; });
  });
}

function rowsFrom_(headers, dataRows) {
  return dataRows.map(function (row) {
    var record = {};
    headers.forEach(function (header, index) {
      if (!header) return;
      var value = row[index];
      if (typeof value === 'string' && /^'[=+\-@]/.test(value)) value = value.slice(1);
      record[header] = value == null ? '' : value;
    });
    return record;
  });
}

function replaceRecords_(records) {
  var headers = [];
  var lookup = {};
  var cleanRecords = collectRecords_(records, headers, lookup);
  if (!cleanRecords.length) throw new Error('The file has no rows to import.');

  var sheet = getSheet_(DATA_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var matrix = matrixFor_(headers, cleanRecords);
  sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
  return matrix.length;
}

function appendRecords_(records) {
  var sheet = getSheet_(DATA_SHEET);
  var headers = existingHeaders_(sheet);
  var lookup = {};
  headers.forEach(function (header, index) {
    if (header) lookup[header.toLowerCase()] = index;
  });

  var cleanRecords = collectRecords_(records, headers, lookup);
  if (!cleanRecords.length) return 0;

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var matrix = matrixFor_(headers, cleanRecords);
  sheet.getRange(sheet.getLastRow() + 1, 1, matrix.length, headers.length).setValues(matrix);
  return matrix.length;
}

function collectRecords_(records, headers, lookup) {
  var cleanRecords = [];
  records.forEach(function (record) {
    if (!record || typeof record !== 'object' || Array.isArray(record)) return;
    var clean = {};
    var hasValue = false;
    Object.keys(record).forEach(function (key) {
      var name = String(key).trim();
      if (!name) return;
      var lower = name.toLowerCase();
      if (!has_(lookup, lower)) {
        headers.push(name);
        lookup[lower] = headers.length - 1;
      }
      var index = lookup[lower];
      var value = record[key];
      if (isBlank_(value)) return;
      if (has_(clean, index)) return;
      clean[index] = value;
      hasValue = true;
    });
    if (hasValue) cleanRecords.push(clean);
  });
  return cleanRecords;
}

function matrixFor_(headers, cleanRecords) {
  return cleanRecords.map(function (clean) {
    return headers.map(function (_, index) {
      return has_(clean, index) ? safeCell_(clean[index]) : '';
    });
  });
}

function safeCell_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : '';
  if (typeof value === 'boolean') return "'" + (value ? 'TRUE' : 'FALSE');
  var text = String(value).trim();
  if (text === '') return '';
  if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(text)) return Number(text);
  return "'" + text;
}

function writeResults_(summary) {
  var sheet = getSheet_(RESULTS_SHEET);
  sheet.clear();

  var rows = [
    ['rowCount', summary.rowCount],
    ['updatedAt', summary.updatedAt],
    [],
    ['DESCRIPTIVES'],
    ['name', 'type', 'n', 'mean', 'median', 'sd', 'distinct']
  ];

  summary.columns.forEach(function (col) {
    rows.push([
      col.name,
      col.type,
      col.n,
      blank_(col.mean),
      blank_(col.median),
      blank_(col.sd),
      col.type === 'word' ? col.distinct : ''
    ]);
  });

  rows.push([]);
  rows.push(['COUNTS']);
  rows.push(['column', 'value', 'count']);
  summary.columns.forEach(function (col) {
    (col.counts || []).forEach(function (item) {
      rows.push([col.name, item.value, item.count]);
    });
  });

  rows.push([]);
  rows.push(['CORRELATIONS']);
  rows.push(['variableA', 'variableB', 'r', 'n']);
  summary.correlations.forEach(function (item) {
    rows.push([item.variableA, item.variableB, blank_(item.r), item.n]);
  });

  var padded = rows.map(function (row) {
    var copy = row.map(function (cell) {
      if (cell === null || cell === undefined) return '';
      var text = String(cell);
      return /^[=+\-@]/.test(text) ? "'" + text : text;
    });
    while (copy.length < RESULTS_WIDTH) copy.push('');
    return copy;
  });

  var range = sheet.getRange(1, 1, padded.length, RESULTS_WIDTH);
  range.setNumberFormat('@');
  range.setValues(padded);
}

function readSummary_() {
  var sheet = getSheet_(RESULTS_SHEET);
  if (sheet.getLastRow() === 0) return null;

  var values = sheet.getDataRange().getDisplayValues().map(function (row) {
    return row.map(function (cell) {
      return /^'[=+\-@]/.test(cell) ? cell.slice(1) : cell;
    });
  });
  var summary = { rowCount: 0, updatedAt: '', columns: [], correlations: [] };
  var section = '';
  var skipNext = false;
  var byName = {};

  for (var i = 0; i < values.length; i++) {
    if (skipNext) {
      skipNext = false;
      continue;
    }
    var row = values[i];
    var label = String(row[0]).trim();

    if (label === 'rowCount' && !section) {
      summary.rowCount = Number(row[1]) || 0;
      continue;
    }
    if (label === 'updatedAt' && !section) {
      summary.updatedAt = String(row[1]);
      continue;
    }
    if (label === 'DESCRIPTIVES' || label === 'COUNTS' || label === 'CORRELATIONS') {
      section = label;
      skipNext = true;
      continue;
    }
    if (!label) continue;

    if (section === 'DESCRIPTIVES') {
      var type = String(row[1]);
      var column = {
        name: label,
        type: type,
        n: Number(row[2]) || 0,
        mean: numberOrNull_(row[3]),
        median: numberOrNull_(row[4]),
        sd: numberOrNull_(row[5])
      };
      if (type === 'word') {
        column.distinct = numberOrNull_(row[6]);
        if (column.distinct === null || column.distinct <= MAX_STORED_COUNTS) column.counts = [];
      }
      summary.columns.push(column);
      byName[column.name] = column;
    } else if (section === 'COUNTS') {
      var target = has_(byName, label) ? byName[label] : null;
      if (target && target.counts) {
        target.counts.push({ value: String(row[1]), count: Number(row[2]) || 0 });
      }
    } else if (section === 'CORRELATIONS') {
      summary.correlations.push({
        variableA: label,
        variableB: String(row[1]),
        r: numberOrNull_(row[2]),
        n: Number(row[3]) || 0
      });
    }
  }

  summary.columns.forEach(function (column) {
    if (column.type === 'word' && column.distinct === null) {
      column.distinct = column.counts ? column.counts.length : 0;
    }
  });

  if (!summary.columns.length && !summary.rowCount) return null;
  return summary;
}

function existingHeaders_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return [];
  return trimHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
}

function trimHeaders_(row) {
  var last = -1;
  for (var i = 0; i < row.length; i++) {
    if (String(row[i]).trim() !== '') last = i;
  }
  var headers = [];
  for (var j = 0; j <= last; j++) headers.push(String(row[j]).trim());
  return headers;
}

function getSheet_(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function has_(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function isBlank_(value) {
  return value === '' || value === null || value === undefined;
}

function firstFilled_(cells) {
  for (var i = 0; i < cells.length; i++) {
    if (!isBlank_(cells[i])) return cells[i];
  }
  return null;
}

function toNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  if (typeof value === 'string' && value.trim() !== '' && isFinite(Number(value))) return Number(value);
  return null;
}

function toWord_(value) {
  if (isBlank_(value)) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MMM d, yyyy');
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  var text = String(value).trim();
  return text === '' ? null : text;
}

function numberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  var number = Number(value);
  return isFinite(number) ? number : null;
}

function mean_(values) {
  var sum = 0;
  for (var i = 0; i < values.length; i++) sum += values[i];
  return sum / values.length;
}

function median_(values) {
  var sorted = values.slice().sort(function (a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) return (sorted[mid - 1] + sorted[mid]) / 2;
  return sorted[mid];
}

function sampleSd_(values) {
  var avg = mean_(values);
  var sum = 0;
  for (var i = 0; i < values.length; i++) sum += Math.pow(values[i] - avg, 2);
  return Math.sqrt(sum / (values.length - 1));
}

function pearson_(xs, ys) {
  var pairsX = [];
  var pairsY = [];
  var length = Math.min(xs.length, ys.length);
  for (var i = 0; i < length; i++) {
    if (xs[i] !== null && ys[i] !== null) {
      pairsX.push(xs[i]);
      pairsY.push(ys[i]);
    }
  }
  if (pairsX.length < 3) return { r: null, n: pairsX.length };

  var meanX = mean_(pairsX);
  var meanY = mean_(pairsY);
  var covariance = 0;
  var sumX = 0;
  var sumY = 0;
  for (var j = 0; j < pairsX.length; j++) {
    var dx = pairsX[j] - meanX;
    var dy = pairsY[j] - meanY;
    covariance += dx * dy;
    sumX += dx * dx;
    sumY += dy * dy;
  }
  var denominator = Math.sqrt(sumX * sumY);
  if (denominator === 0) return { r: null, n: pairsX.length };
  return { r: covariance / denominator, n: pairsX.length };
}

function round6_(value) {
  return Math.round(value * 1000000) / 1000000;
}

function blank_(value) {
  return value === null || value === undefined ? '' : value;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
