/**
 * Google Apps Script — "Export for AI (JSON)" - Namespaced Version
 *
 * This version is designed for integration with existing Apps Script projects.
 * All functionality is contained within the AIExport namespace to avoid conflicts.
 *
 * Usage:
 * 1. Copy this entire code block to your existing Apps Script project
 * 2. Add to your existing onOpen() function:
 *    { name: 'Export for AI (JSON)', functionName: 'AIExport.exportSpreadsheetAsJson' }
 * 3. Or call programmatically: AIExport.exportSpreadsheetAsJson()
 *
 * Requirements: Google Drive API access (automatically requested)
 */

const AIExport = {
  // Internal cache for column letter conversions
  columnLetterCache: [],

  /* ───────────────  MAIN EXPORT FUNCTION  ─────────────── */

  /**
   * Exports the entire spreadsheet's data, formulas, and metadata as JSON to Google Drive.
   * Only non-empty cells and sheets are included to minimize file size.
   */
  exportSpreadsheetAsJson: function() {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    const exportData = {
      explanation:
        'Generated from Google Sheets. Only non-empty cells are included. ' +
        'Use "sheet_name / cell_address" to reference data programmatically.',
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
      const sheetJson = this.buildSheetJson(sheet);
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

    const jsonString = JSON.stringify(exportData, null, 2);
    const base64Data = Utilities.base64Encode(jsonString);
    const filename   = `${ss.getName()}_data.json`;
    const dataUrl    = `data:application/json;base64,${base64Data}`;
    const html       = HtmlService
                         .createHtmlOutput(`<p>Your export is ready: <a href="${dataUrl}" download="${filename}">Download JSON</a></p>`)
                         .setWidth(320)
                         .setHeight(80);

    SpreadsheetApp.getUi().showModalDialog(html, 'Export complete');
  },

  /* ───────────────  HELPER FUNCTIONS  ─────────────── */

  /**
   * Returns sheet-level JSON containing metadata and an array of populated cells.
   * Processes each cell exactly once for efficiency.
   */
  buildSheetJson: function(sheet) {
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
    const mergedLookup = this.buildMergedLookup(range.getMergedRanges());

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
          mergedLookup.has(this.cellA1(r + 1, c + 1, true)) ||
          (rich && rich.getLinkUrl())
        );
        if (!hasData) continue;

        const cell = {
          cell_address: this.cellA1(r + 1, c + 1),
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
          merge_info  : mergedLookup.get(this.cellA1(r + 1, c + 1, true)),
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
  },

  /**
   * Creates a Map<A1, {is_merged, merge_range}> for constant-time lookups.
   */
  buildMergedLookup: function(mergedRanges) {
    const map = new Map();
    mergedRanges.forEach(rg => {
      const mergeA1 = rg.getA1Notation();
      for (let row = rg.getRow(); row < rg.getLastRow() + 1; row++) {
        for (let col = rg.getColumn(); col < rg.getLastColumn() + 1; col++) {
          map.set(this.cellA1(row, col, true), { is_merged: true, merge_range: mergeA1 });
        }
      }
    });
    return map;
  },

  /**
   * Converts 1-based row/col to A1 notation.
   * `raw` flag skips column-letter caching (slightly faster for lookup keys).
   */
  cellA1: function(row, col, raw) {
    if (raw) {
      return this.columnToLetter(col) + row;        // no caching for lookup keys
    }

    // cache the column letter the classic way
    if (!this.columnLetterCache[col]) {
      this.columnLetterCache[col] = this.columnToLetter(col);
    }
    return this.columnLetterCache[col] + row;
  },

  /**
   * Classic column-number → letter(s) conversion.
   */
  columnToLetter: function(n) {
    let s = '';
    while (n) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = (n - m - 1) / 26;
    }
    return s;
  }
};