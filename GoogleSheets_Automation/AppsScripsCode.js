/**
 * Logs or updates tasks from Gmail to a Google Sheet.
 * This script is designed to be easily configurable by editing the CONFIG block below.
 */
function logEmailsToSheet() {
  
    // =================================================================================================
    // --- CONFIGURATION SECTION ---
    // Edit the settings in this block to adapt the script to your needs.
    // =================================================================================================
    const CONFIG = {
      
      // --- 1. Gmail Search Settings ---
  
      // The subject line(s) to search for. The script will look for emails containing ANY of these phrases.
      // Example: ["New task", "Task assigned"]
      search_subjects: ["You have been assigned", "has been assigned to you"],
  
      // Set to true to process ONLY unread emails (recommended for efficiency).
      // Set to false to process all emails (read and unread) that match the criteria.
      search_only_unread: true, 
  
      // Optional: Only process emails received AFTER this date. Use YYYY/MM/DD format.
      // Leave as "" to search all emails regardless of date.
      search_after_date: '2025/04/30',
  
  
      // --- 2. Data Extraction Patterns ---
  
      // The patterns for the main Task ID. The script will look for these letters followed by numbers.
      // This search is case-insensitive (e.g., 'inc' works the same as 'INC').
      // Example: ['INC', 'TASK', 'REQ']
      extract_tarea_id_patterns: ['INC', 'WO', 'TAS'],
  
      // The text labels the script should look for in the email body to find the information.
      extract_titulo_label: "Título:",
      extract_comentario_label: "Descripción detallada:",
      extract_priority_label: "Priority:",
  
  
      // --- 3. Google Sheet Settings ---
  
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
    };
    // =================================================================================================
    // --- END OF CONFIGURATION SECTION ---
    // No need to edit anything below this line.
    // =================================================================================================
  
  
    // --- SCRIPT LOGIC ---
  
    // Build the Gmail search query dynamically from the CONFIG settings
    let queryParts = [];
    if (CONFIG.search_only_unread) {
      queryParts.push('is:unread');
    }
    if (CONFIG.search_subjects && CONFIG.search_subjects.length > 0) {
      queryParts.push(`subject:(${CONFIG.search_subjects.map(s => `"${s}"`).join(' OR ')})`);
    }
    if (CONFIG.search_after_date) {
      queryParts.push(`after:${CONFIG.search_after_date}`);
    }
    const GMAIL_SEARCH_QUERY = queryParts.join(' ');
  
    // Build the Tarea ID regex dynamically
    const idRegex = new RegExp(`(${CONFIG.extract_tarea_id_patterns.join('\\d+|')}\\d+)`, 'i');
  
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG.sheet_name);
  
    if (!sheet) {
      Logger.log(`Sheet "${CONFIG.sheet_name}" not found. SCRIPT TERMINATED.`);
      return;
    }
  
    let headers;
    try {
      headers = sheet.getRange(CONFIG.header_row, 1, 1, sheet.getLastColumn()).getValues()[0];
    } catch (e) {
      Logger.log(`Could not read headers from row ${CONFIG.header_row}. SCRIPT TERMINATED. Error: ${e.toString()}`);
      return;
    }
    
    const col = {
      tarea:      headers.indexOf(CONFIG.column_names.tarea),
      start_date: headers.indexOf(CONFIG.column_names.start_date),
      titulo:     headers.indexOf(CONFIG.column_names.titulo),
      comentario: headers.indexOf(CONFIG.column_names.comentario),
      priority:   headers.indexOf(CONFIG.column_names.priority)
    };
  
    for (const [name, index] of Object.entries(col)) {
      if (index === -1) {
        Logger.log(`Required column "${CONFIG.column_names[name]}" not found in row ${CONFIG.header_row}. Script stopped.`);
        return;
      }
    }
  
    const existingTareas = new Set();
    const lastDataRow = sheet.getLastRow();
    if (lastDataRow > CONFIG.header_row) { 
      const tareaValuesRange = sheet.getRange(CONFIG.header_row + 1, col.tarea + 1, lastDataRow - CONFIG.header_row, 1);
      const tareaValues = tareaValuesRange.getValues();
      for (const row of tareaValues) {
        if (row[0] && row[0].toString().trim() !== "") {
          existingTareas.add(row[0].toString().toUpperCase().trim()); 
        }
      }
    }
    Logger.log(`Found ${existingTareas.size} existing Tarea IDs in the sheet.`);
  
    let threads;
    try {
      threads = GmailApp.search(GMAIL_SEARCH_QUERY);
      Logger.log(`Searching Gmail with query: "${GMAIL_SEARCH_QUERY}". Found ${threads.length} threads.`);
    } catch (e) {
      Logger.log(`Error searching Gmail: ${e.message}. SCRIPT TERMINATED.`);
      return;
    }
    
    const rowsToAdd = []; 
    const messagesToMarkRead = []; 
    let newTareasInThisRun = new Set(); 
  
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (let k = 0; k < messages.length; k++) { 
        const message = messages[k];
        const fullSubject = message.getSubject();
        const receivedDate = message.getDate(); 
        const body = message.getPlainBody(); 
        let tareaValue = ""; 
  
        const match = fullSubject.match(idRegex);
        if (match && match[1]) {
          tareaValue = match[1].toUpperCase().trim(); 
        } else {
          if (CONFIG.search_only_unread && message.isUnread()) messagesToMarkRead.push(message);
          continue; 
        }
        
        if (existingTareas.has(tareaValue) || newTareasInThisRun.has(tareaValue)) {
          Logger.log(`Tarea ID "${tareaValue}" already exists or was processed. Skipping.`);
          if (CONFIG.search_only_unread && message.isUnread()) messagesToMarkRead.push(message);
          continue;
        }
  
        // --- EXTRACTION LOGIC ---
        let tituloValue = "";
        const tituloRegex = new RegExp(`${CONFIG.extract_titulo_label}\\s*([^\\n\\r]*)`, 'i'); 
        const tituloMatch = body.match(tituloRegex);
        if (tituloMatch && tituloMatch[1]) {
            tituloValue = tituloMatch[1].trim();
        }
  
        let comentarioValue = "";
        const comentarioRegex = new RegExp(`${CONFIG.extract_comentario_label}\\s*([\\s\\S]*?)\\s*-`);
        const comentarioMatch = body.match(comentarioRegex);
        if (comentarioMatch && comentarioMatch[1]) {
            comentarioValue = comentarioMatch[1].trim();
        } else {
            const fallbackComentarioRegex = new RegExp(`${CONFIG.extract_comentario_label}\\s*([\\s\\S]*)`, 'i');
            const fallbackMatch = body.match(fallbackComentarioRegex);
            if (fallbackMatch && fallbackMatch[1]) {
                comentarioValue = fallbackMatch[1].trim();
            }
        }
  
        let priorityValue = CONFIG.default_priority;
        const priorityRegex = new RegExp(`${CONFIG.extract_priority_label}\\s*(High|Medium|Low)`, 'i');
        const priorityMatch = body.match(priorityRegex);
        if (priorityMatch && priorityMatch[1]) {
          let foundPriority = priorityMatch[1].toLowerCase();
          priorityValue = foundPriority.charAt(0).toUpperCase() + foundPriority.slice(1);
        }
        // --- END EXTRACTION LOGIC ---
  
        const rowData = [];
        rowData[col.tarea] = tareaValue; 
        rowData[col.start_date] = receivedDate; 
        rowData[col.titulo] = tituloValue;       
        rowData[col.comentario] = comentarioValue; 
        rowData[col.priority] = priorityValue; 
        
        rowsToAdd.push(rowData);
        newTareasInThisRun.add(tareaValue); 
        Logger.log(`New Tarea ID "${tareaValue}" found. Priority: ${priorityValue}. Will be added.`);
  
        if (CONFIG.search_only_unread && message.isUnread()) {
          messagesToMarkRead.push(message);
        }
      }
    }
  
    if (rowsToAdd.length > 0) {
      const numSheetColumns = sheet.getLastColumn(); 
      for (const sparseRowData of rowsToAdd) {
        const completeRowArray = new Array(numSheetColumns).fill(""); 
        for(let i = 0; i < numSheetColumns; i++) { 
          if (sparseRowData[i] !== undefined) { 
            completeRowArray[i] = sparseRowData[i];
          }
        }
        sheet.appendRow(completeRowArray);
      }
      Logger.log(`Added ${rowsToAdd.length} new unique Tarea ID(s) to the sheet "${CONFIG.sheet_name}".`);
    } else {
      Logger.log('No new unique Tarea IDs found to add, or no emails matched criteria with extractable IDs.');
    }
  
    if (messagesToMarkRead.length > 0) {
      const uniqueMessagesToMarkRead = Array.from(new Set(messagesToMarkRead.map(m => m.getId())))
                                       .map(id => messagesToMarkRead.find(m => m.getId() === id));
      GmailApp.markMessagesRead(uniqueMessagesToMarkRead);
      Logger.log(`Marked ${uniqueMessagesToMarkRead.length} message(s) as read.`);
    }
  }
  
  /**
   * Creates a custom menu in the spreadsheet UI to run the script manually.
   */
  function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Custom Email Utilities')
        .addItem('Log Emails to Sheet Now', 'logEmailsToSheet')
        .addToUi();
  }