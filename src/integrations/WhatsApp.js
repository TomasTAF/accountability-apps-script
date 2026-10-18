function configurarWhatsAppWebhook() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = properties.getProperty('WHATSAPP_SPREADSHEET_ID');
  var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Crea o elige una hoja de cálculo y guarda su ID en la propiedad WHATSAPP_SPREADSHEET_ID.');
  }

  Logger.log('Spreadsheet selected: ' + spreadsheet.getName() + ' (' + spreadsheet.getUrl() + ')');
  properties.setProperty('WHATSAPP_SPREADSHEET_ID', spreadsheet.getId());

  var sheet = spreadsheet.getSheetByName('WhatsApp');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('WhatsApp');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Fecha', 'Teléfono', 'Nombre', 'Tipo', 'Contenido', 'ID de mensaje']);
  }
  Logger.log('WhatsApp tab ready in ' + spreadsheet.getName());
}

function doGet(e) {
  var parameters = e && e.parameter ? e.parameter : {};
  var properties = PropertiesService.getScriptProperties();
  var expectedToken = properties.getProperty('WHATSAPP_VERIFY_TOKEN');
  var isValid = parameters['hub.mode'] === 'subscribe' &&
    expectedToken &&
    parameters['hub.verify_token'] === expectedToken;

  return ContentService.createTextOutput(isValid ? parameters['hub.challenge'] : 'Webhook verification failed');
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Webhook request did not include a POST body.');
    }
    var payload = JSON.parse(e.postData.contents);
    console.log('Webhook payload root keys: ' + Object.keys(payload).join(', '));
    var rows = obtenerMensajesWhatsApp_(payload);
    console.log('Extracted message rows: ' + rows.length);
    if (rows.length > 0) {
      var spreadsheetId = PropertiesService.getScriptProperties().getProperty('WHATSAPP_SPREADSHEET_ID');
      if (!spreadsheetId) {
        throw new Error('Ejecuta configurarWhatsAppWebhook antes de recibir mensajes.');
      }

      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheet = spreadsheet.getSheetByName('WhatsApp');
      if (!sheet) {
        throw new Error('No se encontró la pestaña WhatsApp.');
      }
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      console.log('Saved ' + rows.length + ' message row(s) to the WhatsApp sheet.');
    }
  } catch (error) {
    console.error('WhatsApp webhook processing failed: ' + (error.stack || error));
  }

  return ContentService.createTextOutput('EVENT_RECEIVED');
}

function obtenerMensajesWhatsApp_(payload) {
  var rows = [];
  var entries = payload.entry || (payload.field === 'messages' && payload.value
    ? [{ changes: [{ value: payload.value }] }]
    : []);

  entries.forEach(function(entry) {
    (entry.changes || []).forEach(function(change) {
      var value = change.value || {};
      var contacts = value.contacts || [];
      var contactNames = {};
      contacts.forEach(function(contact) {
        contactNames[contact.wa_id] = contact.profile ? contact.profile.name : '';
      });

      (value.messages || []).forEach(function(message) {
        var content = message.text ? message.text.body : JSON.stringify(message[message.type] || '');
        rows.push([
          message.timestamp ? new Date(Number(message.timestamp) * 1000) : '',
          message.from || '',
          contactNames[message.from] || '',
          message.type || '',
          content,
          message.id || ''
        ]);
      });
    });
  });

  return rows;
}