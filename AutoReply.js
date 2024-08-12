function scanDriveAndReply() {
    const config = {
        folderName: 'YOUR_FOLDER_NAME',
        dayLimit: 30,
        senderName: 'Your Name',
        timeZone: Session.getScriptTimeZone()
    };

    const folder = getFolder(config.folderName);
    if (!folder) return;

    const invoiceFiles = extractInvoiceFiles(folder);
    if (invoiceFiles.length === 0) return;

    const searchQuery = buildSearchQuery(invoiceFiles, config.dayLimit, config.timeZone);
    const threads = GmailApp.search(searchQuery);
    if (threads.length === 0) return;

    processThreads(threads, invoiceFiles, config.senderName);
}

// Folder and File Management
function getFolder(folderName) {
    const folders = DriveApp.getFoldersByName(folderName);
    if (!folders.hasNext()) {
        Logger.log(`Folder not found: ${folderName}`);
        return null;
    }
    return folders.next();
}

function extractInvoiceFiles(folder) {
    const invoiceFiles = [];
    const files = folder.getFiles();

    while (files.hasNext()) {
        const file = files.next();
        const matches = file.getName().match(/\d{4}/g);
        if (matches) {
            invoiceFiles.push({ file: file, invoiceNumbers: matches });
        }
    }

    if (invoiceFiles.length === 0) {
        Logger.log('No invoice files found in folder.');
    } else {
        Logger.log(`Extracted invoice numbers: ${invoiceFiles.map(file => file.invoiceNumbers.join(', ')).join('; ')}`);
    }

    return invoiceFiles;
}

// Gmail Processing
function buildSearchQuery(invoiceFiles, dayLimit, timeZone) {
    const invoiceNumbers = invoiceFiles.map(file => file.invoiceNumbers).flat();
    const dateLimit = new Date();
    dateLimit.setDate(dateLimit.getDate() - dayLimit);
    const formattedDate = Utilities.formatDate(dateLimit, timeZone, 'yyyy/MM/dd');

    return `in:inbox subject:(${invoiceNumbers.join(' OR ')}) after:${formattedDate}`;
}

function processThreads(threads, invoiceFiles, senderName) {
    const processedInvoices = getProcessedInvoices();

    threads.forEach(thread => {
        const lastMessage = findLastExternalMessage(thread);
        if (!lastMessage) return;

        const { subject, ccRecipients } = extractMessageDetails(lastMessage);
        const subjectInvoiceNumbers = (subject.match(/\d{4}/g) || []);
        const attachments = findUnprocessedAttachments(subjectInvoiceNumbers, invoiceFiles, processedInvoices);

        if (attachments.length > 0) {
            const unprocessedInvoiceNumbers = subjectInvoiceNumbers.filter(number =>
                !processedInvoices.has(number) && invoiceFiles.some(file => file.invoiceNumbers.includes(number))
            );
            if (unprocessedInvoiceNumbers.length > 0) {
                const replySuccess = replyWithAttachments(lastMessage, subject, ccRecipients, attachments, unprocessedInvoiceNumbers, senderName);
                if (replySuccess) {
                    markInvoicesAsProcessed(unprocessedInvoiceNumbers, thread, processedInvoices);
                }
            }
        } else {
            Logger.log(`No unprocessed invoices or attachments found for invoices: ${subjectInvoiceNumbers.join(', ')}`);
        }
    });

    setProcessedInvoices(processedInvoices);
}

// Message Handling
function findLastExternalMessage(thread) {
    const messages = thread.getMessages();
    for (let j = messages.length - 1; j >= 0; j--) {
        const message = messages[j];
        if (!message.getFrom().includes(Session.getActiveUser().getEmail())) {
            return message;
        }
    }
    Logger.log('No external messages found in thread.');
    return null;
}

function extractMessageDetails(message) {
    return {
        subject: message.getSubject(),
        fromEmail: message.getFrom(),
        ccRecipients: message.getCc()
    };
}

// Invoice and Attachment Management
function findUnprocessedAttachments(subjectInvoiceNumbers, invoiceFiles, processedInvoices) {
    const attachments = [];

    subjectInvoiceNumbers.forEach(subjectNumber => {
        if (!processedInvoices.has(subjectNumber)) {
            invoiceFiles.forEach(invoiceFile => {
                if (invoiceFile.invoiceNumbers.includes(subjectNumber)) {
                    attachments.push(invoiceFile.file);
                }
            });
        }
    });

    // Logger.log(`Attachments found: ${attachments.length}`);
    return attachments;
}

function replyWithAttachments(lastMessage, subject, ccRecipients, attachments, invoicesToProcess, senderName) {
    const replySubject = `Re: ${subject.replace(/^Re: /, '')}`;

    const attachmentNames = attachments.map(file => file.getName().replace(/\.[^/.]+$/, ''));

    // ${invoicesToProcess.join(', ')} = for only invoices that are processed
    const body =
        `Dear Mr./Mrs.,\n\n` +
        `Attached are the payment proof for the invoice(s): ${attachmentNames.join(', ')}\n\n` +
        `Thank you.\n\n` +
        `Best regards,\n${senderName}`;

    try {
        lastMessage.reply(body, {
            attachments: attachments,
            subject: replySubject,
            cc: ccRecipients,
            name: senderName
        });
        return true; // Indicate success
    } catch (error) {
        Logger.log(`Error replying to email: ${error}`);
        return false; // Indicate failure
    }
}

function markInvoicesAsProcessed(invoiceNumbers, thread, processedInvoices) {
    const messageId = thread.getId(); // Get the thread ID

    invoiceNumbers.forEach(number => {
        if (!processedInvoices.has(number)) {
            processedInvoices.set(number, messageId);
            Logger.log(`Invoice number ${number} processed with message ID: ${messageId}`);
        }
    });
}

// Processed Invoices Management
function getProcessedInvoices() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const processedInvoicesString = scriptProperties.getProperty('processedInvoices');
    return processedInvoicesString ? new Map(JSON.parse(processedInvoicesString)) : new Map();
}

function setProcessedInvoices(processedInvoices) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('processedInvoices', JSON.stringify([...processedInvoices]));
}
