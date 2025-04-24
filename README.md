# Google Apps Script Mail Merge for Google Sheets

## Overview
This project provides a user-friendly Mail Merge solution for Google Sheets using Google Apps Script. It allows you to send personalized emails (with optional attachments) to a list of recipients defined in your Google Sheet. The app features a modern popup editor for composing emails, previewing, and sending bulk emails directly from your spreadsheet.

## Features
- Custom menu in Google Sheets for easy access
- Rich email editor (with Quill.js and Tailwind CSS)
- Supports merge fields (e.g., `{{FirstName}}`) for personalized emails
- Preview email before sending
- Send emails with CC, BCC, and attachments
- Configurable sender name and signature
- Error logging and user feedback

## Setup Instructions
1. **Copy the Script:**
   - Open your Google Sheet.
   - Go to **Extensions > Apps Script**.
   - Paste the contents of `Code.gs` into the script editor.
   - Save the project.

2. **Authorize the Script:**
   - On first use, you will be prompted to authorize the script to access your Gmail and Sheets. Follow the on-screen instructions.

3. **Sheet Structure:**
   - Your sheet should have a header row with at least an `Email` column.
   - Optional columns: `CC`, `BCC`, and any other columns you want to use as merge fields (e.g., `FirstName`, `Company`).

   | Email              | FirstName | Company   | CC             | BCC           |
   |--------------------|-----------|-----------|----------------|---------------|
   | user1@email.com    | Alice     | ExampleCo | cc@email.com   | bcc@email.com |
   | user2@email.com    | Bob       | DemoInc   |                |               |

4. **(Optional) Configs Sheet:**
   - Create a sheet named `Configs` for sender name and signature.
   - Example:

   | A           | B                   |
   |-------------|---------------------|
   | SenderName  | Your Name           |
   | Signature   | Best regards,<br>Your Name |

## Usage
1. **Open the Mail Merge Editor:**
   - Reload your Google Sheet. You’ll see a new menu: **Mail Merge**.
   - Click **Mail Merge > Open Mail Merge** to launch the editor.

2. **Compose Your Email:**
   - Enter the subject. Use merge fields like `{{FirstName}}` to personalize.
   - Compose the body using the rich text editor. Merge fields are supported here too.
   - (Optional) Attach files to be sent with every email.

3. **Preview:**
   - Click **Preview First Email** to see how the email will look for the first row.

4. **Send Emails:**
   - Click **Send Emails**. The script will send personalized emails to all rows with a valid `Email`.
   - Success and error messages will be shown. Errors are logged for review.

## Merge Fields
- Use `{{ColumnName}}` in the subject or body to insert values from your sheet.
- Example: `Hello {{FirstName}}, welcome to {{Company}}!`

## Attachments
- You can attach files using the editor. All recipients will receive the same attachments.

## Troubleshooting
- **No "Email" column found:** Ensure your sheet has a header named `Email`.
- **No data in sheet:** Add at least one data row below the header.
- **Authorization errors:** Reopen the editor and re-authorize if prompted.
- **Configs sheet not found:** Sender name and signature are optional. Add a `Configs` sheet as described if needed.
- **Check logs for errors:** Go to **Extensions > Apps Script > View > Logs** for details on failed emails or attachments.

## License
MIT
