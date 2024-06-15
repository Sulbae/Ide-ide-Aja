function tagPeopleOnCondition() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tracker');
  var range = sheet.getDataRange();
  var values = range.getValues();

  var header = values[0];
  var status_column = header.indexOf('Status');
  var email_column = header.indexOf('Email');
  var name_column = header.indexOf('Pemohon')
  var emailTerkirim_column = header.indexOf('Email Terkirim');

  if (status_column === -1 || email_column === -1 || name_column === -1|| emailTerkirim_column === -1) {
    Logger.log('Status, Email, Name, or Email Terkirim column not found.');
    return;
  }

  for (var i = 1; i < values.length; i++) {
    var status = values[i][status_column];
    var emailTerkirim = values[i][emailTerkirim_column];
    
    Logger.log('Row' + (i + 1) + 'status:' + status + ', emailTerkirim:' + emailTerkirim);
    
    if (status && status.trim().toLowerCase() == 'selesai' && !emailTerkirim) {
      var email = values[i][email_column];
      var name = values[i][name_column];
      if (validateEmail(email)) {
        sendEmailNotification(email, name, i + 1);
        sheet.getRange(i + 1, emailTerkirim_column + 1).setValue('Yes');
      } else {
        Logger.log('Invalid email at row' + (i + 1));
      }
    } else {
      Logger.log('No email sent for row' + (i + 1) + 'because status is not "Selesai" or email already sent.');
    }
  }
}

function sendEmailNotification(email, name, row) {
  var subject = 'Admin Campus: Permohonan Dokumen Selesai';
  var body = 'Halo ' + name + ',\n\n' + 'Proses dokumen kamu sudah selesai, bisa diambil segera.\n\n' + 'Mohon abaikan pesan ini jika kamu sudah ambil!\n\n'+ 'Salam,\nAdmin Campus';
  MailApp.sendEmail(email, subject, body)
}

function validateEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() == 'Tracker') {
    tagPeopleOnCondition();
  }
}
