// Add custom menu to Google Sheet
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Mail Merge')
        .addItem('Open Mail Merge', 'showPopup')
        .addToUi();
}

// Show popup with email editor
function showPopup() {
    const html = HtmlService.createHtmlOutput(getPopupHtml())
        .setTitle('Mail Merge Editor')
        .setWidth(650)
        .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, 'Mail Merge Editor');
}

// HTML for popup
function getPopupHtml() {
    return `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <!-- Load Tailwind CSS from CDN -->
          <script src="https://cdn.tailwindcss.com"></script>
          <!-- Load Quill.js from CDN -->
          <link href="https://cdn.jsdelivr.net/npm/quill@1.3.7/dist/quill.snow.css" rel="stylesheet">
          <script src="https://cdn.jsdelivr.net/npm/quill@1.3.7/dist/quill.min.js"></script>
          <style>
            body { font-family: 'Inter', sans-serif; }
            .ql-toolbar { border-top-left-radius: 0.375rem; border-top-right-radius: 0.375rem; }
            .ql-container { border-bottom-left-radius: 0.375rem; border-bottom-right-radius: 0.375rem; }
            .ql-editor { min-height: 150px; }
            #preview { max-height: 200px; overflow-y: auto; }
          </style>
        </head>
        <body class="bg-gray-100 p-6">
          <div class="max-w-2xl mx-auto bg-white rounded-lg shadow-lg p-6">
            <h3 class="text-2xl font-semibold text-gray-800 mb-6">Mail Merge Editor</h3>
            <div class="mb-4">
              <label class="block text-sm font-medium text-gray-700 mb-1">Subject</label>
              <input id="subject" type="text" class="w-full px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" placeholder="Enter email subject">
            </div>
            <div class="mb-4">
              <label class="block text-sm font-medium text-gray-700 mb-1">Body (use {{column_name}} for merge fields)</label>
              <div id="editor" class="bg-white rounded-md"></div>
            </div>
            <div class="mb-4">
              <label class="block text-sm font-medium text-gray-700 mb-1">Attachments</label>
              <input id="attachments" type="file" multiple class="w-full px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500">
            </div>
            <div class="flex space-x-4 mb-6">
              <button onclick="previewEmail()" class="px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 transition duration-200">Preview First Email</button>
              <button onclick="sendEmails()" class="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition duration-200">Send Emails</button>
            </div>
            <div id="preview" class="p-4 bg-gray-50 border border-gray-200 rounded-md"></div>
          </div>
  
          <script>
            // Initialize Quill editor for body
            const quill = new Quill('#editor', {
              theme: 'snow',
              modules: {
                toolbar: [
                  ['bold', 'italic', 'underline'],
                  ['link', 'blockquote'],
                  [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                  [{ 'color': [] }, { 'background': [] }],
                  ['clean']
                ]
              }
            });
  
            // Sanitize text for safe HTML rendering
            function sanitizeText(text) {
              const div = document.createElement('div');
              div.textContent = text;
              return div.innerHTML;
            }
  
            // Get selected file names
            function getAttachmentNames() {
              const files = document.getElementById('attachments').files;
              return Array.from(files).map(file => file.name).join(', ') || 'None';
            }
  
            // Preview email for first row
            function previewEmail() {
              const subject = document.getElementById('subject').value;
              const body = quill.root.innerHTML;
              const attachmentNames = getAttachmentNames();
  
              google.script.run
                .withSuccessHandler(result => {
                  document.getElementById('preview').innerHTML = 
                    '<div class="text-sm text-gray-700"><strong>CC:</strong> ' + (result.cc || 'None') + '</div>' +
                    '<div class="text-sm text-gray-700"><strong>BCC:</strong> ' + (result.bcc || 'None') + '</div>' +
                    '<div class="text-sm text-gray-700"><strong>Attachments:</strong> ' + (result.attachments || 'None') + '</div>' +
                    '<div class="text-sm text-gray-700"><strong>Subject:</strong> ' + result.subject + '</div>' +
                    '<div class="text-sm text-gray-700 mt-2"><strong>Body:</strong><br>' + result.body + '</div>';
                })
                .getPreviewEmail(subject, body);
            }
  
            // Send emails with uploaded attachments
            function sendEmails() {
              const subject = document.getElementById('subject').value;
              const body = quill.root.innerHTML;
              const files = document.getElementById('attachments').files;
  
              if (files.length > 0) {
                const filePromises = Array.from(files).map(file => {
                  return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve({ name: file.name, content: reader.result.split(',')[1] });
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                  });
                });
  
                Promise.all(filePromises)
                  .then(fileData => {
                    google.script.run
                      .withSuccessHandler(() => {
                        alert('Emails sent successfully!');
                        google.script.host.close();
                      })
                      .withFailureHandler(err => alert('Error: ' + err))
                      .sendMailMerge(subject, body, fileData);
                  })
                  .catch(err => alert('Error reading files: ' + err));
              } else {
                google.script.run
                  .withSuccessHandler(() => {
                    alert('Emails sent successfully!');
                    google.script.host.close();
                  })
                  .withFailureHandler(err => alert('Error: ' + err))
                  .sendMailMerge(subject, body, []);
              }
            }
          </script>
        </body>
      </html>
    `;
}

// Get sheet data efficiently
function getSheetData() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 2) return { headers: [], data: [] };

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
    return { headers, data };
}

// Get config data from Configs sheet
function getConfig() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet;
    try {
        configSheet = spreadsheet.getSheetByName('Configs');
        if (!configSheet) throw new Error('Configs sheet not found');
    } catch (e) {
        throw new Error('Error accessing Configs sheet: ' + e.message);
    }

    const lastRow = configSheet.getLastRow();
    if (lastRow < 1) return { senderName: '', signature: '' };

    const configData = configSheet.getRange(1, 1, lastRow, 2).getValues();
    const config = { senderName: '', signature: '' };

    configData.forEach(row => {
        if (row[0] === 'SenderName') config.senderName = row[1] || '';
        if (row[0] === 'Signature') config.signature = row[1] || '';
    });

    return config;
}

// Preview email for first row
function getPreviewEmail(subject, body) {
    const { headers, data } = getSheetData();
    const config = getConfig();
    if (data.length === 0) throw new Error('No data in sheet');

    const row = data[0];
    const columnIndices = {
        cc: headers.indexOf('CC'),
        bcc: headers.indexOf('BCC')
    };

    const mergedSubject = replaceMergeFields(subject, headers, row);
    const mergedBody = replaceMergeFields(body, headers, row) + `<b/><b/>${config.signature}`;
    const cc = columnIndices.cc !== -1 ? row[columnIndices.cc] : '';
    const bcc = columnIndices.bcc !== -1 ? row[columnIndices.bcc] : '';

    return {
        cc,
        bcc,
        subject: mergedSubject,
        body: mergedBody
    };
}

// Replace merge fields efficiently
function replaceMergeFields(text, headers, row) {
    let result = text;
    for (let i = 0; i < headers.length; i++) {
        const placeholder = `{{${headers[i]}}}`;
        const regex = new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        result = result.replace(regex, row[i] || '');
    }
    return result;
}

// Send mail merge with uploaded attachments
function sendMailMerge(subject, body, fileData) {
    const { headers, data } = getSheetData();
    const config = getConfig();
    const emailColumn = headers.indexOf('Email');
    if (emailColumn === -1) throw new Error('No "Email" column found in sheet');

    const columnIndices = {
        cc: headers.indexOf('CC'),
        bcc: headers.indexOf('BCC')
    };

    let emailCount = 0;
    const errorLog = [];

    // Prepare attachments from uploaded files
    const attachments = fileData.map(file => {
        try {
            const blob = Utilities.newBlob(Utilities.base64Decode(file.content), MimeType.PDF, file.name);
            return blob;
        } catch (e) {
            errorLog.push(`Failed to process attachment ${file.name}: ${e.message}`);
            return null;
        }
    }).filter(blob => blob !== null);

    for (const row of data) {
        if (!row[emailColumn]) continue;

        try {
            const mergedSubject = replaceMergeFields(subject, headers, row);
            const mergedBody = replaceMergeFields(body, headers, row) + `<b/><b/>${config.signature}`;
            const cc = columnIndices.cc !== -1 ? row[columnIndices.cc] : '';
            const bcc = columnIndices.bcc !== -1 ? row[columnIndices.bcc] : '';

            // Prepare email options
            const emailOptions = {
                to: row[emailColumn],
                subject: mergedSubject,
                htmlBody: mergedBody,
            };
            if (cc) emailOptions.cc = cc;
            if (bcc) emailOptions.bcc = bcc;
            if (config.senderName) emailOptions.name = config.senderName;
            if (attachments.length > 0) emailOptions.attachments = attachments;

            // Send email using GmailApp
            GmailApp.sendEmail(row[emailColumn], mergedSubject, '', emailOptions);
            emailCount++;
        } catch (e) {
            errorLog.push(`Failed to send email to ${row[emailColumn]}: ${e.message}`);
        }
    }

    // Log errors if any
    if (errorLog.length > 0) {
        Logger.log('Errors during mail merge:\n' + errorLog.join('\n'));
    }

    SpreadsheetApp.getUi().alert(`Sent ${emailCount} emails successfully!${errorLog.length > 0 ? ' Check logs for errors.' : ''}`);
}