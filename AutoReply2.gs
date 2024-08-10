// Improved version that handle several 4 digits number in the email and label it with only the numbers that are present google drive folder, which allow unlabeled numbers to be processed and send with attachments in the futuer

function replyWithPaymentProof() {
  const folderName = "Payment Proof";
  const processedLabelName = "Payment Proof Sent";
  const daysAgo = 30; // Timeframe to scan emails

  Logger.log("Starting script...");

  // Get the payment folder and label
  const folder = getFolderByName(folderName);
  if (!folder) {
    Logger.log("Folder not found: " + folderName);
    return;
  }

  const label = getOrCreateLabel(processedLabelName);

  // Step 1: Collect all invoice numbers from filenames
  const invoiceToFilesMap = collectInvoiceFiles(folder);

  if (Object.keys(invoiceToFilesMap).length === 0) {
    Logger.log("No valid invoice numbers found in any file.");
    return;
  }

  Logger.log(`Collected ${Object.keys(invoiceToFilesMap).length} invoice numbers from files.`);

  // Build search query for all collected memo numbers
  const searchQuery = buildSearchQuery(Object.keys(invoiceToFilesMap), daysAgo);

  Logger.log(`Search query: ${searchQuery}`);

  // Step 2: Fetch and process emails based on the search query
  const threads = GmailApp.search(searchQuery);

  Logger.log(`Found ${threads.length} threads with search query.`);

  if (threads.length === 0) {
    Logger.log("No threads found with the search query.");
    return;
  }

  processThreads(threads, invoiceToFilesMap, label);

  Logger.log("Script completed.");
}

function buildSearchQuery(invoiceNumbers, daysAgo) {
  const dateQuery = `after:${formatDate(new Date(new Date().setDate(new Date().getDate() - daysAgo)))}`;
  const memoQueries = invoiceNumbers.map(num => `"${num}"`).join(' OR ');
  return `${dateQuery} (${memoQueries})`;
}

function processThreads(threads, invoiceToFilesMap, label) {
  threads.forEach(thread => {
    const messages = thread.getMessages();
    const latestMessage = messages[messages.length - 1];
    const subject = latestMessage.getSubject();
    const matchedInvoiceNumbers = matchInvoicesToSubject(subject, invoiceToFilesMap);

    if (matchedInvoiceNumbers.length > 0) {
      // Track already processed memo numbers for this thread
      const processedMemos = getProcessedMemos(thread, label);
      const newMemos = matchedInvoiceNumbers.filter(memo => !processedMemos.includes(memo));

      if (newMemos.length > 0) {
        sendReplyWithAttachments(thread, latestMessage, newMemos, invoiceToFilesMap, label);
        updateProcessedMemos(thread, processedMemos.concat(newMemos), label);
      } else {
        Logger.log(`Thread with subject "${subject}" has already processed all memos.`);
      }
    } else {
      Logger.log(`No matching invoices found for subject "${subject}".`);
    }
  });
}

function sendReplyWithAttachments(thread, message, invoiceNumbers, invoiceToFilesMap, label) {
  try {
    const firstName = getFirstName(message.getTo());
    const uniqueFilesForThread = Array.from(new Set(invoiceNumbers.flatMap(invoiceNumber => invoiceToFilesMap[invoiceNumber])));

    if (uniqueFilesForThread.length === 0) {
      Logger.log("No files found for the matched invoice numbers.");
      return;
    }

    const replyBody = createReplyBody(firstName, uniqueFilesForThread.map(file => file.getName()).join(', '));

    // Prepare attachments array
    const attachments = uniqueFilesForThread.map(file => file.getBlob());

    // Truncate subject if it exceeds 200 characters
    let replySubject = thread.getFirstMessageSubject();
    if (replySubject.length > 200) {
      replySubject = replySubject.substring(0, 197) + "...";
    }

    // Send the reply directly within the thread
    message.reply(replyBody, {
      attachments: attachments,
      cc: message.getCc(),
      subject: replySubject,
      name: "Your Name"  // You can specify your name or leave it as default
    });

    Logger.log(`Replied to the sender with subject: "${replySubject}" and included CC recipients with attachments: ${uniqueFilesForThread.map(file => file.getName()).join(', ')}`);

  } catch (error) {
    Logger.log("Failed to send reply: " + error.message);
  }
}

function getProcessedMemos(thread, label) {
  const labels = thread.getLabels();
  const processedMemos = [];
  labels.forEach(l => {
    const match = l.getName().match(/\b\d{4}\b/g);
    if (match) {
      processedMemos.push(...match);
    }
  });
  return Array.from(new Set(processedMemos)); // Return unique memo numbers
}

function updateProcessedMemos(thread, allMemos, label) {
  // Remove existing label with memo numbers
  const existingLabel = thread.getLabels().find(l => l.getName().includes(label.getName()));
  if (existingLabel) {
    thread.removeLabel(existingLabel);
  }

  // Label the thread with all processed memo numbers
  const newLabel = GmailApp.createLabel(`${label.getName()} ${allMemos.join(' ')}`);
  thread.addLabel(newLabel);
}

function formatDate(date) {
  return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

function getFolderByName(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

function collectInvoiceFiles(folder) {
  const invoiceToFilesMap = {};
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const invoiceNumbers = extractMemoNumbersFromFilename(file.getName());

    invoiceNumbers.forEach(invoiceNumber => {
      if (!invoiceToFilesMap[invoiceNumber]) {
        invoiceToFilesMap[invoiceNumber] = [];
      }
      invoiceToFilesMap[invoiceNumber].push(file);
    });
  }

  return invoiceToFilesMap;
}

function matchInvoicesToSubject(subject, invoiceToFilesMap) {
  const matchedInvoices = [];

  Object.keys(invoiceToFilesMap).forEach(invoiceNumber => {
    if (subject.includes(invoiceNumber)) {
      matchedInvoices.push(invoiceNumber);
    }
  });

  return matchedInvoices;
}

function getFirstName(email) {
  const firstEmail = email.split(',')[0].trim();
  const localPart = firstEmail.split('@')[0];
  const nameParts = localPart.split(/[.\-_]/);
  const firstName = nameParts[0];
  return firstName.charAt(0).toUpperCase() + firstName.slice(1);
}

function createReplyBody(firstName, fileNames) {
  const salutation = `Dear Mr./Mrs.`;
  const body = `Attached are the payment proof documents for the following memos:\n${fileNames}.`;
  const closing = `Thank you\n\nBest regards,\n[Your Name]`;

  return `${salutation}\n\n${body}\n\n${closing}`;
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}

function extractMemoNumbersFromFilename(filename) {
  const match = filename.match(/\b\d{4}\b/g);
  return match ? match : [];
}
