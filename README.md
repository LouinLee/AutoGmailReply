# Drive and Gmail Integration Script

## Overview

This script automates the process of finding and replying to emails related to specific invoice numbers. It scans a Google Drive folder for invoice files, extracts invoice numbers from the file names, searches for emails with subjects containing these invoice numbers, and replies to those emails with the relevant attachments.

## Features

- Scans a specified Google Drive folder for files with invoice numbers in their names.
- Searches Gmail inbox for emails with subjects containing the extracted invoice numbers.
- Replies to emails with attachments of the relevant invoices.
- Marks processed emails and invoices to avoid duplicate processing.

## Setup

### 1. Create and Configure the Script

1. **Open Google Apps Script:**
   - Go to [Google Apps Script](https://script.google.com/) and create a new project.

2. **Copy the Script:**
   - Copy the entire script provided in this document and paste it into the script editor.

3. **Update Folder Name:**
   - Replace `'YOUR_FOLDER_NAME'` with the name of the Google Drive folder you want to scan.

4. **Customize Reply Content:**
   - Update the `body` of the reply message and the sender's name in the `scanDriveAndReply` function.

### 2. Set Up Permissions

The script requires permissions to access your Google Drive and Gmail. The first time you run the script, it will prompt you to authorize these permissions. Follow the on-screen instructions to grant access.

## Usage

### Running the Script

1. **Run Manually:**
   - To execute the script manually, click the run button (`▶️`) in the Google Apps Script editor.

2. **Set Up Triggers (Optional):**
   - You can set up a time-driven trigger to run the script automatically at specified intervals. To do this:
     - Go to `Triggers` (clock icon) in the Google Apps Script editor.
     - Click `Add Trigger` and configure the time-based trigger according to your needs.

### Important Notes

- **Invoice Number Format:** The script assumes that invoice numbers are four-digit numbers (e.g., `1234`) present in the filenames and email subjects.
- **Processed Emails:** Emails are marked as processed using labels in Gmail. Ensure that labels are used consistently to avoid confusion.
- **Attachments Handling:** The script searches for `.pdf` files that match invoice numbers. If filenames contain multiple invoice numbers, they will be included as attachments.

## Functions

### `scanDriveAndReply()`

Scans the specified Google Drive folder for invoice files, searches for related emails in Gmail, and replies with the relevant attachments.

### `getProcessedMessageIds()`

Retrieves a set of processed message IDs from script properties.

### `setProcessedMessageIds(processedIds)`

Updates the list of processed message IDs in script properties.

### `markAsProcessed(invoiceNumbers, thread)`

Labels the email thread with labels corresponding to processed invoice numbers.

## Logging

The script uses `Logger.log` for logging various actions and statuses. You can view the logs in the Apps Script editor by clicking `View` > `Logs`.

## Contact

For any issues or questions, please contact [louin.liman@gmail.com]

---
