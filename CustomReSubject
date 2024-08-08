function replyWithPaymentProof() {
  var folderName = "Payment Proof";
  var processedLabelName = "Payment Proof Sent";
  var daysAgo = 30; // Adjust this to your desired timeframe

  Logger.log("Starting script...");

  var threads = GmailApp.search('subject:"memo"');
  Logger.log("Found " + threads.length + " threads.");

  var folder = getFolderByName(folderName);
  if (!folder) {
    Logger.log("Folder not found: " + folderName);
    return;
  }
  
  var label = getOrCreateLabel(processedLabelName);

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysAgo);

  var invoiceThreads = {}; // Object to keep track of the latest thread per invoice number

  threads.forEach(function(thread) {
    var firstMessageDate = thread.getMessages()[0].getDate();
    if (firstMessageDate > cutoffDate) {
      var subject = thread.getMessages()[0].getSubject();
      var invoiceNumbers = extractInvoiceNumbers(subject);

      if (invoiceNumbers.length > 0) {
        invoiceNumbers.forEach(function(invoiceNumber) {
          if (!invoiceThreads[invoiceNumber] || thread.getLastMessageDate() > invoiceThreads[invoiceNumber].getLastMessageDate()) {
            invoiceThreads[invoiceNumber] = thread;
          }
        });
      } else {
        Logger.log("No valid invoice numbers found in subject: " + subject);
      }
    } else {
      Logger.log("Skipped old email with subject: " + thread.getMessages()[0].getSubject());
    }
  });

  var filesToSend = [];
  var invoiceToFilesMap = {}; // Map to track files for each invoice number

  // Collect all invoice numbers from filenames and corresponding threads
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var invoiceNumbersFromFile = extractInvoiceNumbersFromFilename(fileName);

    if (invoiceNumbersFromFile.length > 0) {
      invoiceNumbersFromFile.forEach(function(invoiceNumber) {
        if (invoiceThreads[invoiceNumber]) {
          if (!invoiceToFilesMap[invoiceNumber]) {
            invoiceToFilesMap[invoiceNumber] = [];
          }
          invoiceToFilesMap[invoiceNumber].push(file);
        }
      });
    }
  }

  // Create a map to associate threads with the files to send
  var threadFilesMap = {};

  for (var invoiceNumber in invoiceToFilesMap) {
    var thread = invoiceThreads[invoiceNumber];
    if (!threadFilesMap[thread.getId()]) {
      threadFilesMap[thread.getId()] = [];
    }
    var filesForThread = invoiceToFilesMap[invoiceNumber];
    threadFilesMap[thread.getId()].push(...filesForThread);
  }

  // Send replies with the appropriate attachments
  for (var threadId in threadFilesMap) {
    var thread = GmailApp.getThreadById(threadId);
    if (thread && !threadHasLabel(thread, label)) {
      var filesForThread = threadFilesMap[threadId];
      var uniqueFilesForThread = [...new Set(filesForThread)]; // Ensure unique attachments

      var message = thread.getMessages()[0];
      var firstName = getFirstName(message.getTo());
      var replyBody = createReplyBody(firstName, uniqueFilesForThread.map(file => file.getName()).join(', '));
      var replyOptions = { htmlBody: replyBody, attachments: uniqueFilesForThread.map(file => file.getBlob()) };

      // Create a new email with "Re: " in the subject
      var originalSubject = message.getSubject();
      var replySubject = `Re: ${originalSubject}`;

      GmailApp.sendEmail({
        to: message.getFrom(),
        cc: message.getCc(),
        bcc: message.getBcc(),
        subject: replySubject,
        htmlBody: replyBody,
        attachments: uniqueFilesForThread.map(file => file.getBlob())
      });

      Logger.log("Replied with attachments: " + uniqueFilesForThread.map(file => file.getName()).join(', '));
      
      thread.addLabel(label);
    } else {
      Logger.log("Email already processed or not found for invoice numbers: " + Object.keys(invoiceToFilesMap).join(', '));
    }
  }

  Logger.log("Script completed.");
}
