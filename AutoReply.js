function scanDriveAndReply() {
  var folderName = 'YOUR_FOLDER_NAME';
  var folders = DriveApp.getFoldersByName(folderName);

  if (!folders.hasNext()) {
    Logger.log('Folder not found: ' + folderName);
    return;
  }

  var folder = folders.next();
  var files = folder.getFiles();
  var invoiceNumbers = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var matches = fileName.match(/\d{4}/g);
    if (matches) {
      invoiceNumbers.push(...matches);
    }
  }

  Logger.log('Extracted invoice numbers: ' + invoiceNumbers.join(', '));

  var invoiceString = invoiceNumbers.join(' OR ');
  var threads = GmailApp.search('in:inbox subject:(' + invoiceString + ')');

  if (threads.length > 0) {
    Logger.log('Found ' + threads.length + ' threads.');

    // Get processed message IDs
    var processedMessageIds = getProcessedMessageIds();

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      var lastMessage = null;

      // Find the latest message not sent by us
      for (var j = messages.length - 1; j >= 0; j--) {
        var message = messages[j];
        var fromEmail = message.getFrom();
        if (!fromEmail.includes(Session.getActiveUser().getEmail())) {
          lastMessage = message;
          break;
        }
      }

      if (!lastMessage) {
        Logger.log('No external messages found in thread.');
        continue;
      }

      var subject = lastMessage.getSubject();
      var fromEmail = lastMessage.getFrom();
      var ccRecipients = lastMessage.getCc(); // Get the CC recipients

      Logger.log('Processing thread with subject: ' + subject);
      // Logger.log('From email: ' + fromEmail);
      // Logger.log('CC recipients: ' + ccRecipients);

      var matchedInvoices = subject.match(/\d{4}/g) || [];
      var invoicesToProcess = matchedInvoices.filter(function (number) {
        return invoiceNumbers.includes(number) && !processedMessageIds.has(number);
      });

      Logger.log('Matched invoices: ' + invoicesToProcess.join(', '));

      if (invoicesToProcess.length > 0) {
        var attachments = [];
        // Loop through each invoice number to find matching files
        invoicesToProcess.forEach(function (number) {
          var filesInFolder = folder.getFilesByName(number + '.pdf');
          while (filesInFolder.hasNext()) {
            var file = filesInFolder.next();
            attachments.push(file);
          }
        });

        // Handle cases where filenames may contain multiple invoice numbers
        var allFiles = folder.getFiles();
        while (allFiles.hasNext()) {
          var file = allFiles.next();
          var fileName = file.getName();
          if (invoicesToProcess.some(number => fileName.includes(number))) {
            attachments.push(file);
          }
        }

        Logger.log('Attachments found: ' + attachments.length);

        if (attachments.length > 0) {
          var replySubject = 'Re: ' + subject.replace(/^Re: /, '');
          var body = "Dear Mr./Mrs.,\n\n" +
            "Attached are the payment proof for the invoice(s):\n" +
            invoicesToProcess.join(', ') + "\n\n" +
            "Thank you.\n\n" +
            "Best regards,\n" +
            "Your Name"; // Replace 'John Doe' with your actual name

          try {
            lastMessage.reply(body, {
              attachments: attachments,
              subject: replySubject,
              cc: ccRecipients // Include CC recipients in the reply
            });

            // Mark invoices as processed
            invoicesToProcess.forEach(function (invoice) {
              processedMessageIds.add(invoice);
            });
            processedMessageIds.add(lastMessage.getId()); // Mark message ID as processed

            markAsProcessed(invoicesToProcess, thread);
            Logger.log('Replied with invoices: ' + invoicesToProcess.join(', '));
            break;
          } catch (error) {
            Logger.log('Error replying to email: ' + error);
          }
        } else {
          Logger.log('No attachments found for invoices: ' + invoicesToProcess.join(', '));
        }
      } else {
        Logger.log('No matching invoices to process.');
      }
    }

    // Save updated processed message IDs
    setProcessedMessageIds(processedMessageIds);
  } else {
    Logger.log('No matching threads found.');
  }
}

function getProcessedMessageIds() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const processedIdsString = scriptProperties.getProperty('processedMessageIds');
  return processedIdsString ? new Set(processedIdsString.split(',')) : new Set();
}

function setProcessedMessageIds(processedIds) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('processedMessageIds', Array.from(processedIds).join(','));
}

function markAsProcessed(invoiceNumbers, thread) {
  invoiceNumbers.forEach(function (number) {
    var labelName = 'Processed: ' + number;
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }
    thread.addLabel(label);
    Logger.log('Marked as processed with label: ' + labelName);
  });
}
