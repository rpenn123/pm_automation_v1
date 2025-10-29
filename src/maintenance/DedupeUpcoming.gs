function dedupeUpcomingBySfidOrNameLoc() {
  const ss = SpreadsheetApp.getActive();
  const CFG = CONFIG;
  const UP = CFG.UPCOMING_COLS;
  const sheet = ss.getSheetByName(CFG.SHEETS.UPCOMING);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getMaxColumns();
  if (lastRow < 3) return;

  const tsCol = getHeaderColumnIndex(sheet, CFG.LAST_EDIT.AT_HEADER); // "Last Edit At (hidden)"
  const vals = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Build key + timestamp for each row
  const rows = vals.map((row, i) => {
    const sfid = String(row[UP.SFID - 1] || "").trim();
    const name = normalizeString(String(row[UP.PROJECT_NAME - 1] || "").trim());
    const loc  = normalizeString(String(row[UP.LOCATION - 1] || "").trim());
    const ts   = tsCol > 0 ? row[tsCol - 1] : null; // raw Date or blank
    const key  = sfid ? `SFID:${sfid}` : `NAMELOC:${name}|${loc}`;
    return { idx: i + 2, key, ts: ts instanceof Date ? ts.getTime() : 0 };
  });

  // Group by key
  const groups = new Map();
  for (const r of rows) {
    if (!groups.has(r.key)) groups.set(r.key, []);
    groups.get(r.key).push(r);
  }

  // Decide deletions: keep the newest timestamp; if tie/blank, keep bottom-most
  const toDelete = [];
  for (const list of groups.values()) {
    if (list.length <= 1) continue;
    list.sort((a, b) => {
      if (b.ts !== a.ts) return b.ts - a.ts; // newest ts first
      return a.idx - b.idx; // tie: later rows considered newer → keep last; delete earlier
    });
    // Keep list[0], delete the rest
    for (let i = 1; i < list.length; i++) toDelete.push(list[i].idx);
  }

  toDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
}