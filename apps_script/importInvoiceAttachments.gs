/**** CONFIG ****/
const SHEET_NAME = 'Invoice Database';
const GMAIL_LABEL = 'Invoices';
const DRIVE_FOLDER_NAME = 'Invoice PDFs';  // will be created if missing

/**** MAIN ENTRY POINT: IMPORT ATTACHMENTS ****/
function importInvoiceAttachments() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  // Column H stores Gmail message IDs to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds =
    lastRow > 1 ? sheet.getRange(2, 8, lastRow - 1, 1).getValues().flat() : [];

  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  if (!label) return Logger.log('Label "' + GMAIL_LABEL + '" not found');

  const folder = getOrCreateFolder_(DRIVE_FOLDER_NAME);
  const threads = label.getThreads(0, 50); // recent labelled threads

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (existingIds.includes(msgId)) return; // already processed

      const attachments = message.getAttachments({ includeInlineImages: false });
      if (!attachments.length) return;

      attachments.forEach(att => {
        if (!att.getContentType().match(/pdf/i)) return;

        const file = folder.createFile(att);
        const fileName = file.getName();

        const parsed = parseInvoiceFileName_(fileName);
        if (!parsed) return Logger.log('Could not parse: ' + fileName);

        const { invoiceId, customer, country, partsType, amount, status, dateStr } = parsed;

        sheet.appendRow([
          invoiceId,
          customer,
          country,
          partsType,
          amount,
          status,
          dateStr,
          msgId
        ]);
      });
    });
  });
}

/**** HELPERS ****/

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

/**
 * Expected file format:
 * INV-1001_AirTech-GmbH_Germany_Bearings_320.50_Paid_2024-01-05.pdf
 */
function parseInvoiceFileName_(fileName) {
  const base = fileName.replace(/\.pdf$/i, '');
  const parts = base.split('_');
  if (parts.length < 7) return null;

  return {
    invoiceId: parts[0],
    customer: parts[1].replace(/-/g, ' '),
    country: parts[2],
    partsType: parts[3],
    amount: parseFloat(parts[4]),
    status: parts[5],
    dateStr: parts[6]
  };
}
