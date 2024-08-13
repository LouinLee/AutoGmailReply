# GDrive and Gmail Integration Script

## Overview

This is a Google Apps Script designed to automate the process of scanning a specific Google Drive folder for invoice files, searching Gmail for related email threads, and replying with the relevant invoices attached. This script ease the process of responding to clients or stakeholders with necessary payment proof documents.

## Features

- Scans a specified Google Drive folder for invoice files.
- Extracts invoice numbers from file names.
- Searches Gmail for threads with subjects matching the invoice numbers.
- Replies to the most recent external email in each thread with the corresponding invoice files attached.
- Tracks processed invoices to avoid duplicate replies.

## Configuration

Modify the following configuration parameters in the `scanDriveAndReply` function:

- `folderName`: The name of the folder in Google Drive containing the invoice files.
- `dayLimit`: The number of days to look back in Gmail for relevant email threads.
- `senderName`: Your name, which will appear as the sender in the reply emails.
- `timeZone`: The time zone for date formatting.

## Usage

1. Update the `config` object with your folder name, day limit, sender name, and time zone.
2. Save the script.
3. Run the script function manually or set up a time-based trigger to run it automatically at regular intervals (optional).

### Folder and File Management

- `getFolder(folderName)`: Retrieves the folder object by its name.
- `extractInvoiceFiles(folder)`: Extracts invoice numbers from file names in the specified folder.

### Gmail Processing

- `buildSearchQuery(invoiceFiles, dayLimit, timeZone)`: Builds a Gmail search query to find relevant email threads.
- `processThreads(threads, invoiceFiles, senderName)`: Processes email threads and replies with the corresponding attachments.

### Message Handling

- `findLastExternalMessage(thread)`: Finds the most recent external message in a thread.
- `extractMessageDetails(message)`: Extracts details such as subject and CC recipients from a message.

### Invoice and Attachment Management

- `findUnprocessedAttachments(subjectInvoiceNumbers, invoiceFiles, processedInvoices)`: Finds unprocessed attachments for the given invoice numbers.
- `replyWithAttachments(lastMessage, subject, ccRecipients, attachments, invoicesToProcess, senderName)`: Replies to the email with the specified attachments.

### Processed Invoices Management

- `getProcessedInvoices()`: Retrieves the set of processed invoices from script properties.
- `setProcessedInvoices(processedInvoices)`: Saves the set of processed invoices to script properties.

## Contact

For any issues or questions, please contact [sirlinjj@gmail.com]

---
