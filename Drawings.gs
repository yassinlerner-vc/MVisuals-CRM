// ============================================================
//  DRAWINGS — GENERATE ID
// ============================================================
function generateDrawingId() {
  const sheet = getSheet('Drawings');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[DRW.DRAWING_ID] || '').match(/^DRW(\d+)$/);
    if (match) { const num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
  });
  return 'DRW' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  DRAWINGS — LIST
// ============================================================
function getDrawings() {
  const sheet = getSheet('Drawings');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  return data.slice(1)
    .filter(row => row[DRW.DRAWING_ID])
    .map(row => ({
      id:         String(row[DRW.DRAWING_ID]),
      personId:   row[DRW.PERSON_ID],
      personName: row[DRW.PERSON_NAME],
      amount:     Number(row[DRW.AMOUNT] || 0),
      currency:   row[DRW.CURRENCY],
      date:       row[DRW.DATE] ? Utilities.formatDate(new Date(row[DRW.DATE]), tz, 'yyyy-MM-dd') : '',
      method:     row[DRW.METHOD],
      notes:      row[DRW.NOTES],
      createdAt:  row[DRW.CREATED_AT]
        ? Utilities.formatDate(new Date(row[DRW.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
    }))
    .sort((a, b) => b.date.localeCompare(a.date));
}

// ============================================================
//  DRAWINGS — CREATE
//  Always 100% to one partner, always drawn from the pool —
//  no split table, no PaidBy field needed.
// ============================================================
function createDrawing(formData) {
  const sheet = getSheet('Drawings');
  if (!sheet) return { success: false, error: 'Drawings sheet missing.' };

  const amount = Number(formData.amount) || 0;
  if (amount <= 0) return { success: false, error: 'Enter a valid amount.' };
  if (!formData.personId) return { success: false, error: 'Select a partner.' };
  if (!formData.date) return { success: false, error: 'Select a date.' };

  const drawingId = generateDrawingId();
  const now  = new Date();
  const user = Session.getActiveUser().getEmail();

  sheet.appendRow([
    drawingId, formData.personId, formData.personName,
    amount, formData.currency, formData.date,
    formData.method || '', formData.notes || '', user, now
  ]);

  writeLog('Drawings_Log', 'Drawings', drawingId,
    drawingId + ' — ' + formData.personName, '', '',
    'Amount', '', amount, 'Drawing recorded by ' + user);

  return { success: true, drawingId };
}

// ============================================================
//  DRAWINGS — DELETE
// ============================================================
function deleteDrawing(drawingId) {
  const sheet = getSheet('Drawings');
  if (!sheet) return { success: false, error: 'Drawings sheet missing.' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][DRW.DRAWING_ID]) === String(drawingId)) {
      const personName = data[i][DRW.PERSON_NAME];
      sheet.deleteRow(i + 1);
      writeLog('Drawings_Log', 'Drawings', drawingId,
        drawingId + ' — ' + personName, '', '',
        'Status', 'Recorded', 'Deleted', 'Drawing deleted');
      return { success: true };
    }
  }
  return { success: false, error: 'Drawing not found.' };
}
