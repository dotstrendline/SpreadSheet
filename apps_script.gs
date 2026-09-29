// Paste this whole file into: Google Sheet -> Extensions -> Apps Script
// Then change TOKEN below to your own password (same value goes in SHEET_TOKEN).

const TOKEN = 'CHANGE_THIS_TO_YOUR_OWN_PASSWORD';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out_({ ok: false, error: 'bad token' });

    const values = d.values || [];
    if (!values.length) return out_({ ok: false, error: 'no values' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(d.title);
    const created = !sh;
    if (created) sh = ss.insertSheet(d.title);

    // make rectangular
    const rows = values.length;
    const cols = Math.max.apply(null, values.map(r => r.length));
    const grid = values.map(r => r.concat(Array(cols - r.length).fill('')));

    // make sure the tab is big enough
    if (sh.getMaxRows() < rows + 20) sh.insertRowsAfter(sh.getMaxRows(), rows + 20 - sh.getMaxRows());
    if (sh.getMaxColumns() < cols + 5) sh.insertColumnsAfter(sh.getMaxColumns(), cols + 5 - sh.getMaxColumns());

    // keep "09:29" / "2026-09-29 09:29:10" as plain text (Sheets would convert them)
    sh.getRange(1, 1, 1, cols).setNumberFormat('@');
    if (d.text_first_col) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');

    sh.clearContents();
    sh.getRange(1, 1, rows, cols).setValues(grid);

    if (created) {
      sh.getRange(1, 1, 1, cols).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (d.freeze_col) sh.setFrozenColumns(1);
      if (d.color_scores) {
        const rng = sh.getRange(2, 2, sh.getMaxRows() - 1, sh.getMaxColumns() - 1);
        const pos = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0).setFontColor('#178a3f').setBold(true).setRanges([rng]).build();
        const neg = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0).setFontColor('#d92626').setBold(true).setRanges([rng]).build();
        sh.setConditionalFormatRules([pos, neg]);
      }
    }
    return out_({ ok: true, rows: rows, cols: cols });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
