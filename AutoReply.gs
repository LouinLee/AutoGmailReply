function replyWithPaymentProof() {
  var folderName = "Payment Proof";
  var processedLabelName = "Payment Proof Sent";
  var daysAgo = 14; // Adjust this to your desired timeframe

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
      var replyOptions = { attachments: uniqueFilesForThread.map(file => file.getBlob()) };

      message.replyAll(replyBody, replyOptions);
      Logger.log("Replied with attachments: " + uniqueFilesForThread.map(file => file.getName()).join(', '));
      
      thread.addLabel(label);
    } else {
      Logger.log("Email already processed or not found for invoice numbers: " + Object.keys(invoiceToFilesMap).join(', '));
    }
  }

  Logger.log("Script completed.");
}

function getFolderByName(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function getOrCreateLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  return label ? label : GmailApp.createLabel(name);
}

function extractInvoiceNumbers(subject) {
  var matches = subject.match(/memo\s*[\/\s\-]*([\d\s]+)(?=\s|$)/i);
  if (matches) {
    return matches[1].trim().split(/\s+/).filter(num => num.length === 4);
  }
  return [];
}

function extractInvoiceNumbersFromFilename(filename) {
  var match = filename.match(/^(\d{4}(?:\s\d{4})*)/);
  if (match) {
    return match[1].trim().split(/\s+/).filter(num => num.length === 4);
  }
  return [];
}

function getFirstName(email) {
  return email.split('@')[0].split('.').map(name => name.charAt(0).toUpperCase() + name.slice(1)).join(' ');
}

function createReplyBody(firstName, fileNames) {
  // Helper function to remove any file extension from a filename
  function removeFileExtension(fileName) {
    return fileName.replace(/\.[^/.]+$/, ''); // Regex to remove any extension
  }

  // Process file names to remove extensions
  const processedFileNames = fileNames
    .split(', ')
    .map(fileName => removeFileExtension(fileName))
    .join(', ');

  const salutation = `Dear Mr./Mrs. ${firstName},`;
  const body = `Attached are the payment proof documents for the following memos:\n${processedFileNames}.`;
  const closing = `Thank you\n\nBest regards,\n[Your Name]`;

  return `${salutation}\n\n${body}\n\n${closing}`;
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}
