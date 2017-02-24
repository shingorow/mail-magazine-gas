(function(global) {

  var MailMagazine = function(config, isTest) {
    if (! config && config.email) {
      Browser.msgBox('config を設定して下さい');
      return;
    }
    var ssConfig,
        emailConfig,
        recipients,
        emailTemplate,
        recipient,
        emailMessage,
        emailLabel;
    

    ssConfig = config.spreadsheet;
    if (! ssConfig.ssId) throw('スプレッドシートの ID を config に設定して下さい');
    emailLabel = ssConfig.emailLabel || 'メールアドレス';
    emailConfig = config.email || {};
    ssConfig.sheetName = isTest ? ssConfig.testSheet : ssConfig.sheetName;
    recipients = getRecipientsData(ssConfig);
    emailTemplate = getEmailTemplate(ssConfig);
    
    for (var i = 0, length = recipients.length; i < length; i++) {
      recipient = recipients[i];
      if (recipient[emailLabel]) {
        emailMessage = createMessage(emailTemplate, recipient);
        sendEmail(recipient[emailLabel], emailMessage, emailConfig);
      }
    }


    function　getValuesFromSpreadsheet(ssId, sheetName, displayValues) {
      var ss,
          sheet,
          range,
          values;

      ss = SpreadsheetApp.openById(ssId);
      sheet = ss.getSheetByName(sheetName);
      range = sheet.getDataRange();
      values = displayValues ? range.getDisplayValues() : range.getValues();
      
      return values;
    }
    

    function　getRecipientsData(ssConfig) {
      var values = [],
          row = [],
          value,
          keys = [],
          key,
          tmp = {},
          list = [],
          ssId,
          sheetName,
          displayValues;
      ssId = ssConfig.ssId;
      sheetName = ssConfig.sheetName;
      displayValues = ssConfig.displayValues;
      values = getValuesFromSpreadsheet(ssId, sheetName, displayValues);
      keys = values.shift();
      
      for (var i = 0; i < values.length; i++) {
        row = values[i];
        for (var j = 0; j < keys.length; j++) {
          key = keys[j];
          value = row[j];
          tmp[key] = value;
        }
        
        list.push(tmp);
        tmp = {};
      }
      
      return list;
    }
    

    function　getEmailTemplate(ssConfig) {
      var contents = {},
          row,
          ssId,
          messageSheet,
          values,
          displayValue;
      ssId = ssConfig.ssId;
      messageSheet = ssConfig.messageSheet || 'メッセージ';
      displayValue = ssConfig.displayValue;
      
      values = getValuesFromSpreadsheet(ssId, messageSheet, displayValue);
      for (var i = 0; i < values.length; i++) {
        row = values[i];
        contents[row[0]] = row[1];
      }
      
      return contents;
    }
    

    function　createMessage(emailTemplate, recipientData) {
      var emailContents,
          subject,
          re;
      emailContents = JSON.parse(JSON.stringify(emailTemplate));
      for (var i in emailContents) {
        for (var j in recipientData) {
          re = new RegExp('{{' + j + '}}', 'g');
          emailContents[i] = emailContents[i].replace(re, recipientData[j]);
        }
      }

      return emailContents;
    }
    

    function sendEmail(recipient, emailMessage, emailConfig) {
      var options = {},
          subject,
          body;
      if(emailConfig.attachments) options.attachments = emailConfig.attachments;
      if(emailConfig.cc) options.cc = emailConfig.cc;
      if(emailConfig.bcc) options.bcc = emailConfig.bcc;
      if(emailConfig.from) options.from = emailConfig.from;
      if(emailMessage['本文:HTML']) options.htmlBody = emailMessage['本文:HTML'];
      if(emailConfig.inlineImages) options.inlineImages = emailConfig.inlineImages;
      if(emailConfig.name) options.name = emailConfig.name;
      if(emailConfig.replyTo) options.replyTo = emailConfig.replyTo;
      
      body = emailMessage['本文:テキスト'];
      subject = emailMessage['件名'];
      //Logger.log('To %s Body: %s', recipient, body);
      GmailApp.sendEmail(recipient, subject, body, options)
    }
  };
  
  global.MailMagazine = MailMagazine;

}(this));


/**
 * スプレッドシートのメールアドレスリストにメールを送信する
 *
 * @param {Object} config 使用するスプレッドシートやメールの設定を含むコンフィグ
 */
function send(config) {
  var isTest = false;
  MailMagazine(config, isTest);
}

function test(config) {
  var isTest = true;
  MailMagazine(config, isTest);

}
