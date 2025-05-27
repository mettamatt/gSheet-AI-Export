/**
 * Google Apps Script — “Export for AI (JSON)”
 *
 * Usage: the custom menu **Export Tools → Export for AI (JSON)** appears
 * whenever the file is opened.
 */

/* ───────────────  UI HOOK  ─────────────── */

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu('Export Tools', [{ name: 'Export for AI (JSON)', functionName: 'exportSpreadsheetAsJson' }]);
}

/* ───────────────  ENTRY POINT  ─────────────── */

function exportSpreadsheetAsJson() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const exportData = {
    explanation:
      'Generated from Google Sheets. Only non-empty cells are included. ' +
      'Use “sheet_name / cell_address” to reference data programmatically.',
    spreadsheet_metadata: {
      spreadsheet_name: ss.getName(),
      export_timestamp: new Date().toISOString(),
      total_sheets : 0,
      total_rows   : 0,
      total_columns: 0,
      total_cells  : 0,
      total_formulas: 0,
    },
    sheets: [],
  };

  sheets.forEach(sheet => {
    const sheetJson = buildSheetJson(sheet);
    if (sheetJson.cells.length) {              // keep only non-empty sheets
      exportData.sheets.push(sheetJson);

      // roll up totals
      const m = exportData.spreadsheet_metadata;
      m.total_sheets   += 1;
      m.total_rows     += sheetJson.sheet_metadata.num_rows;
      m.total_columns  =  Math.max(m.total_columns, sheetJson.sheet_metadata.num_columns);
      m.total_cells    += sheetJson.sheet_metadata.num_cells;
      m.total_formulas += sheetJson.sheet_metadata.num_formulas;
    }
  });

  const blob   = Utilities.newBlob(JSON.stringify(exportData, null, 2),
                                   MimeType.JSON,
                                   `${ss.getName()}_data.json`);
  const file   = DriveApp.getFileById(ss.getId()).getParents().next().createFile(blob); // same folder
  const html   = HtmlService
                   .createHtmlOutput(`<p>Your export is ready: <a href="${file.getUrl()}" target="_blank">Download JSON</a></p>`)
                   .setWidth(320)
                   .setHeight(80);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export complete');
}

/* ───────────────  HELPERS  ─────────────── */

/**
 * Returns sheet-level JSON containing metadata and an array of populated cells.
 * Processes each cell exactly once for efficiency.
 */
function buildSheetJson(sheet) {
  const range      = sheet.getDataRange();
  const numRows    = range.getNumRows();
  const numCols    = range.getNumColumns();

  // Bulk fetches (single Sheet API call each)
  const values         = range.getValues();
  const formulas       = range.getFormulas();
  const comments       = range.getComments();
  const notes          = range.getNotes();
  const validations    = range.getDataValidations();
  const richTextValues = range.getRichTextValues();

  // Build fast lookup for merged-range membership
  const mergedLookup = buildMergedLookup(range.getMergedRanges());

  // Iterate cells once
  const cells       = [];
  let   formulaCnt  = 0;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {

      const value    = values[r][c];
      const formula  = formulas[r][c];
      const comment  = comments[r][c];
      const note     = notes[r][c];
      const dv       = validations[r][c];
      const rich     = richTextValues[r][c];

      if (formula) formulaCnt++;

      // Skip empty cells fast
      const hasData = (
        value !== '' && value !== null ||
        (formula && formula.trim())     ||
        (comment && comment.trim())     ||
        (note && note.trim())           ||
        dv                              ||
        mergedLookup.has(cellA1(r + 1, c + 1, true)) ||
        (rich && rich.getLinkUrl())
      );
      if (!hasData) continue;

      const cell = {
        cell_address: cellA1(r + 1, c + 1),
        value       : value,
        formula     : formula || undefined,
        comment     : comment || undefined,
        note        : note || undefined,
        data_validation: dv ? {
          criteria_type  : dv.getCriteriaType().toString(),
          criteria_values: dv.getCriteriaValues(),
          allow_invalid  : dv.getAllowInvalid(),
          strict         : dv.getStrict()
        } : undefined,
        merge_info  : mergedLookup.get(cellA1(r + 1, c + 1, true)),
        hyperlink   : rich ? rich.getLinkUrl() : undefined,
      };

      // Remove undefined keys for compactness
      Object.keys(cell).forEach(k => cell[k] === undefined && delete cell[k]);
      cells.push(cell);
    }
  }

  return {
    sheet_name : sheet.getName(),
    sheet_id   : sheet.getSheetId(),
    sheet_index: sheet.getIndex(),
    sheet_metadata: {
      num_rows    : numRows,
      num_columns : numCols,
      num_cells   : numRows * numCols,
      num_formulas: formulaCnt,
      visibility  : sheet.isSheetHidden() ? 'Hidden' : 'Visible',
    },
    cells
  };
}

/**
 * Creates a Map<A1, {isMerged, mergeRange}> for constant-time lookups.
 */
function buildMergedLookup(mergedRanges) {
  const map = new Map();
  mergedRanges.forEach(rg => {
    const mergeA1 = rg.getA1Notation();
    for (let row = rg.getRow(); row < rg.getLastRow() + 1; row++) {
      for (let col = rg.getColumn(); col < rg.getLastColumn() + 1; col++) {
        map.set(cellA1(row, col, true), { is_merged: true, merge_range: mergeA1 });
      }
    }
  });
  return map;
}

/**
 * Converts 1-based row/col to A1 notation.
 * `raw` flag skips column-letter caching (slightly faster for lookup keys).
 */
const columnLetterCache = [];
function cellA1(row, col, raw) {
  return (raw ? columnToLetter(col) : (columnLetterCache[col] ||= columnToLetter(col))) + row;
}

/** Classic column-number → letter(s) */
function columnToLetter(n) {
  let s = '';
  while (n) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}
