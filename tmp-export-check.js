const fs = require('fs');
function buildWorkbookBlob(entries) {
  const rows = buildWorksheetRows(entries);
  const sharedStrings = buildSharedStrings(rows);
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
      name: "xl/sharedStrings.xml",
      content: createSharedStringsXml(sharedStrings)
    },
    {
      name: "xl/worksheets/sheet1.xml",
      content: createWorksheetXml(rows, sharedStrings)
    }
  ];

  const zipBytes = buildZipArchive(files);
  return new Blob([zipBytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function buildWorksheetRows(entries) {
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

  return [headers].concat(entries.map(entry => [
    entry.startDate,
    entry.endDate,
    String(diffInDays(entry.startDate, entry.endDate) + 1),
    entry.flow,
    entry.mood,
    Array.isArray(entry.symptoms) ? entry.symptoms.join(", ") : "",
    entry.notes || "",
    entry.id
  ]));
}

function buildSharedStrings(rows) {
  const values = [];
  const indexByValue = new Map();

  rows.forEach(row => {
    row.forEach(value => {
      const stringValue = String(value || "");
      if (!indexByValue.has(stringValue)) {
        indexByValue.set(stringValue, values.length);
        values.push(stringValue);
      }
    });
  });

  return {
    values,
    indexByValue,
    count: rows.reduce((total, row) => total + row.length, 0)
  };
}

function createWorksheetXml(rows, sharedStrings) {
  const sheetRows = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndexValue) => {
      const cellReference = `${columnName(columnIndexValue)}${rowIndex + 1}`;
      const stringValue = String(value || "");
      const sharedIndex = sharedStrings.indexByValue.get(stringValue) || 0;
      return `<c r="${cellReference}" t="s"><v>${sharedIndex}</v></c>`;
    }).join("");

    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");

  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>",
    "<sheetFormatPr defaultRowHeight=\"15\"/>",
    `<dimension ref="A1:H${rows.length}"/>`,
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

function createSharedStringsXml(sharedStrings) {
  const items = sharedStrings.values
    .map(value => `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`)
    .join("");

  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.count}" uniqueCount="${sharedStrings.values.length}">`,
    items,
    "</sst>"
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
    "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>",
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
    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
    "</Relationships>"
  ].join("");
}
function createStylesXml() {
  return [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>",
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>",
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
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
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


function diffInDays(startDate, endDate) {
  const start = parseDateValue(startDate);
  const end = parseDateValue(endDate);
  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  return Math.round((end - start) / millisecondsPerDay);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}


(async () => { const blob = buildWorkbookBlob([{ id: 'entry-1', startDate: '2026-03-12', endDate: '2026-03-15', flow: 'Medium', mood: 'Balanced', symptoms: ['Cramps'], notes: 'Test note' }]); const buffer = Buffer.from(await blob.arrayBuffer()); fs.writeFileSync('D:\\Period Tracker\\artifacts-test-export.xlsx', buffer); })().catch(err => { console.error(err); process.exit(1); });