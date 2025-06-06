# How to Use the Google Sheet Automation Script
This guide explains how to set up and configure the Google Apps Script to automate logging tasks from your Gmail to your Google Sheet.

## How to Set Up and Configure the Script
You only need to edit the configuration section at the very top of the script file. All the logic will adapt to your settings automatically.

### 1. Open the Script Editor

In your Google Sheet, go to the menu Extensions > Apps Script.

### 2. Locate the CONFIG Block

At the very top of the script file, you will see a section marked CONFIGURATION SECTION. You will make all your changes here.

### 3. Update the Settings

Change the values inside the CONFIG block to match your specific needs. Each setting is explained below:
Gmail Search Settings
```js
// The subject line(s) to search for. The script will look for emails containing ANY of these phrases.
 // Example: ["New task", "Task assigned"]
 search_subjects: ["You have been assigned", "has been assigned to you"],

 // Set to true to process ONLY unread emails (recommended for efficiency).
 // Set to false to process all emails (read and unread) that match the criteria.
 search_only_unread: true, 

 // Optional: Only process emails received AFTER this date. Use 'YYYY/MM/DD' format.
 // Leave as "" to search all emails regardless of date.
 search_after_date: '2025/04/30',
 ```


* search_subjects: Change the text in quotes to match the subject lines of the emails you want the script to find. You can have one or more phrases.

* search_only_unread: This is a simple switch. Set it to true to only look at unread emails (this is much faster and recommended). Set it to false if you need the script to look at all emails every time.
search_after_date: Use this if you only want to process emails from a certain date onwards. If you want to search all emails, just set it to "".

#### Data Extraction Patterns

```js
// The patterns for the main Task ID. The script will look for these letters followed by numbers.
 // This search is case-insensitive (e.g., 'inc' works the same as 'INC').
 // Example: ['INC', 'TASK', 'REQ']
 extract_tarea_id_patterns: ['INC', 'WO', 'TAS'],

 // The text labels the script should look for in the email body to find the information.
 extract_titulo_label: "Título:",
 extract_comentario_label: "Descripción detallada:",
 extract_priority_label: "Priority:",
```
* extract_tarea_id_patterns: Change INC, WO, TAS to the prefixes of your own Task IDs (e.g., REQ, JOB, TICKET).
  
* extract_..._label: Update these with the exact text labels used in your emails to identify the Title, Comment, and Priority information.

#### Google Sheet Settings

```js
// The name of the sheet (tab) where your data is stored.
 sheet_name: 'Remedy',

 // The row number where your column titles (headers) are located.
 header_row: 1,

 // The exact titles of your columns in the sheet.
 column_names: {
   tarea:      "TAREA",
   start_date: "START",
   titulo:     "TITULO",
   comentario: "Comentario",
   priority:   "PRIORITY"
 },

 // The default value to use for the Priority column if it's not found in the email.
 // This should be one of the options in your dropdown selector.
 default_priority: "BACKLOG"
```
* sheet_name: Change 'Remedy' to the name of your sheet tab.
header_row: Set this to the row number where your titles are (e.g., 1 for the first row, 2 for the second).

* column_names: Make sure the names in quotes (like "TAREA", "START") perfectly match your column titles in the spreadsheet.
default_priority: Set the default value you want for the priority if it can't be found in an email.

### 4. Save the Script

After making your changes, click the Save project icon (looks like a floppy disk) at the top of the editor.

## Running the Script

### First-Time Manual Run (Required)

You must run the script manually once to grant it the necessary permissions to access your Gmail and Google Sheets.

1. In the Apps Script editor, make sure the function logEmailsToSheet is selected in the dropdown menu at the top.

2. Click the Run button.

3. A pop-up will appear asking for authorization. Follow the prompts, choose your Google account, and click Allow.
4. After running it once, a new menu called "Custom Email Utilities" will appear in your spreadsheet. You can use this to run the script manually at any time.

### Setting Up Automatic Triggers

To make the script run automatically on a schedule:

1. In the Apps Script editor, click on the Triggers icon (looks like an alarm clock) on the left sidebar.
2. Click the + Add Trigger button in the bottom-right.
3. Configure the trigger settings. For example, to run every day at 8 AM:
4. Choose which function to run: logEmailsToSheet
5. Select event source: Time-driven
6. Select type of time based trigger: Day timer
7. Select time of day: 8am - 9am
8. Click Save. You can add multiple triggers for different times of the day.