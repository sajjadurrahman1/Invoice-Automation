const SHEET_NAME = 'Invoice Database';
const GMAIL_LABEL = 'Invoices';

/**** BACKFILL GMAIL IDS FOR EXISTING ROWS ****/
function fillGmailIdColumn() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return Logger.log('No data rows found');

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 8);
  const data = dataRange.getValues();

  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  if (!label) return Logger.log('Label "' + GMAIL_LABEL + '" not found');

  const threads = label.getThreads(0, 500);
  const idByInvoice = {};

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const subj = message.getSubject();
      const match = subj.match(/INV[-\s]?(\d+)/i);
      if (match) {
        const invoiceId = ('INV-' + match[1]).toUpperCase();
        idByInvoice[invoiceId] = message.getId();
      }
    });
  });

  for (let i = 0; i < data.length; i++) {
    const invoiceId = String(data[i][0]).trim().toUpperCase();
    if (!invoiceId) continue;
    if (data[i][7]) continue; // already filled

    if (idByInvoice[invoiceId]) {
      data[i][7] = idByInvoice[invoiceId];
    }
  }

  dataRange.setValues(data);
}
