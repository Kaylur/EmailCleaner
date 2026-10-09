# Email Cleaner for Microsoft Word

A Word VBA macro that opens a pop-up window to clean raw email data, extract usernames, and insert the results into a document as bulleted lists.

## What it does

Raw email data often includes extra text such as match scores, labels, or brackets. This tool removes that extra text and leaves only the email address. It can also pull the username (the part before the @) from each address.

**Example input**

```
jsmith42@example.com (100%)
mary.jones@example.org (85%)
<JSMITH42@example.com> (72%)
mailto:test.user@example.net
```

**Cleaned emails**

```
jsmith42@example.com
mary.jones@example.org
test.user@example.net
```

**Usernames**

```
jsmith42
mary.jones
test.user
```

## Features

- Removes percentages, brackets, labels, and other extra text
- Handles one or many emails per line
- Converts addresses to lowercase and removes duplicates
- Extracts usernames with a separate button
- Inserts emails and usernames as separate bulleted lists anywhere in the document
- Stays open while you work, so you can click into different sections of a report
- Each insert can be undone with a single Ctrl+Z

## Requirements

- Microsoft Word for Windows (desktop version)
- Macros enabled

## Installation

1. Download `EmailCleaner.bas`.
2. Open Word and press **Alt+F11** to open the VBA editor.
3. In the project list, select **Normal**, then go to **File > Import File** and choose `EmailCleaner.bas`.
4. In Word, go to **File > Options > Trust Center > Trust Center Settings > Macro Settings** and check **Trust access to the VBA project object model**.
5. Run the `BuildEmailCleanerForm` macro once. This creates the pop-up window.
6. You can turn off the trust setting from step 4 after the form is built.

Installing into Normal makes the tool available in every document.

## Usage

1. Run the `EmailCleaner` macro. You can add it to the Quick Access Toolbar for one-click access.
2. Paste the raw email data into the top box.
3. Click **Clean Emails** to fill the cleaned email list.
4. Click **Get Usernames** to fill the username list.
5. Click in the document where the email list should go, then click **Insert Emails at Cursor**.
6. Click in the section where the username list should go, then click **Insert Usernames at Cursor**.

If the cursor is on a line that already has text, such as a section heading, the list is inserted on the line below it. Lists use Word's built-in List Bullet style.

You can edit the contents of either result box before inserting.

## Buttons

| Button | Action |
| --- | --- |
| Clean Emails | Extracts and cleans email addresses from the input box |
| Get Usernames | Extracts the username from each email address |
| Clear All | Clears all boxes |
| Close | Closes the pop-up |
| Insert Emails at Cursor | Inserts the cleaned emails as a bulleted list |
| Insert Usernames at Cursor | Inserts the usernames as a bulleted list |

## Limitations

- Entries that are not in a standard email format are skipped. Review the results against the source data before finalizing a report.
- Addresses are converted to lowercase. If you need to keep the original capitalization, edit the result box before inserting.
- The form must be rebuilt if it is deleted. To rebuild, remove `frmEmailCleaner` in the VBA editor and run `BuildEmailCleanerForm` again.

## Troubleshooting

**"Word is blocking access to the VBA project"**
Turn on **Trust access to the VBA project object model** (see Installation, step 4) and run `BuildEmailCleanerForm` again.

**"The Email Cleaner form has not been built yet"**
Run `BuildEmailCleanerForm` once before using `EmailCleaner`.

**Macros are disabled**
Check your macro settings in the Trust Center or contact your IT administrator.
