/**
 * @OnlyCurrentDoc
 * Indicates that this automation will only attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the limited scope.
 */

/**
 * Adds custom menus to the Google Sheets UI upon opening the spreadsheet.
 * 
 * This function is automatically triggered when the spreadsheet is opened and 
 * creates custom menus titled 'Drive utilities' and 'Gmail utilities'.
 * 
 * @param {Event} e - The event object containing information about the opening event (optional).
 * 
 * @return {void} This function does not return a value.
 * 
 * @example
 * onOpen();
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂️ Drive utilities 🗂️')
    .addItem('📑 Create named files', 'copyFiles')
    .addItem('📂 Create named subfolders', 'copyFolders')
    .addSeparator()
    .addItem('🔗 Retrieve File links', 'retrieveFiles')
    .addItem('🔗 Retrieve Subfolder links', 'retrieveFolders')
    .addSeparator()
    .addItem('📑 Create Copies', 'createCopies')
    .addToUi();
  ui.createMenu('📧 Gmail utilities 📧')
      .addItem('📧 Send mail merge', 'sendEmails')
      .addItem('Display Sheet Names','displaySheetNames')
      .addToUi();
}

/**
 * Displays a prompt dialog to the user with the names of all available sheets
 * in the active spreadsheet and allows them to select one.
 *
 * This function retrieves the names of all sheets, presents them to the user in a
 * prompt dialog, and processes the user's selection. If the user selects a valid
 * sheet name, further actions can be taken based on the selected sheet. If the
 * user cancels the prompt or selects an invalid name, appropriate alerts are shown.
 *
 * @function displaySheetNames
 * @returns {void} This function does not return a value.
 */
function displaySheetNames() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let sheetNames = sheets.map(sheet => sheet.getName());
  console.log(`sheetNames = '${sheetNames.join(', ')}'`);

  // Present the sheet names to the user using a dialog
  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt('Choose a sheet', 'Available sheets: ' + sheetNames.join(', '), ui.ButtonSet.OK_CANCEL);
  console.log(`User input: '${result.getResponseText()}'`);

  // Continue processing based on the user's input
  if (result.getSelectedButton() == ui.Button.OK) {
    let selectedSheetName = result.getResponseText();
    if (sheetNames.includes(selectedSheetName)) {
      let selectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(selectedSheetName);
      // Continue processing with selectedSheet
    } else {
      ui.alert('The selected sheet name is not valid.');
      console.warn(`The selected sheet name is not valid.\rUser input: '${result.getResponseText()}'`);
    }
  } else {
    ui.alert('Process cancelled.');
  }
}

/**
 * Prompts the user for input and returns their response as a string.
 * 
 * Displays a prompt dialog to the user; if canceled, returns an empty string.
 * 
 * @param {string} promptMessage - The message to display in the prompt dialog.
 * @return {string} The user's input from the prompt; an empty string if canceled.
 * 
 * @example
 * const userInput = getUserInput("Please enter your name:");
 * Logger.log(`User input: ${userInput}`);
 */
function getUserInput(promptMessage) {
  return SpreadsheetApp.getUi().prompt(promptMessage).getResponseText();
}

/**
 * Fills a template string with values from the provided data object.
 * 
 * This function replaces all occurrences of {{key}} in the template string 
 * with corresponding values from the data object. If a key does not exist in 
 * the object, it will be replaced with an empty string.
 * 
 * @param {string} template The template string containing {{}} markers for replacement.
 * @param {object} data An object with key-value pairs to replace the {{}} markers.
 * @returns {string} The filled template string with the values from the data object.
 *
 * @example
 * const result = fillInTemplateFromObject_('Hello, {{name}}!', { name: 'Alice' });
 * console.log(result); // Output: 'Hello, Alice!'
 */
function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);

  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    const text = data[key.replace(/[{}]+/g, "")] || "";
    return RegExp.escape(text.replace(/\n/g, '<br>'));
  });
  return JSON.parse(template_string);
}

/**
 * Checks if the provided header row is valid.
 * 
 * This function verifies that the header row contains at least two unique,
 * non-empty headers. It returns true if the conditions are met, otherwise false.
 *
 * @param {Array<string>} headers An array of header strings to validate.
 * @returns {boolean} Returns true if the header row is valid; otherwise false.
 */
function isValidHeaderRow(headers) {
  const uniqueHeaders = new Set(headers.filter(header => header.trim() !== ""));
  return uniqueHeaders.size >= 2;
}

/**
 * Converts a 2D array into an array of objects, mapping column headers to cell values.
 *
 * @param {Array<Array<string>>} data A 2D array where the first row contains column headers.
 * @param {string} urlColumn The column index (1-based) of the sheet where the file URLs will be written
 * @param {string} nameHeader The header name for filenames to copy
 * @param {string} urlHeader The header name to receive new copy URLs
 * @returns {Array<Object>} An array of objects representing rows with key-value pairs
 * based on column headers and their corresponding cell values.
 *
 * @example
 * const data = [
 *   ['Name', 'Age', 'City'],
 *   ['Alice', '30', 'New York'],
 *   ['Bob', '25', 'San Francisco']
 * ];
 *
 * const result = convertToObjects(data);
 * console.log(result);
 * // Output:
 * // [
 * //   { Name: 'Alice', Age: '30', City: 'New York' },
 * //   { Name: 'Bob', Age: '25', City: 'San Francisco' }
 * // ]
 */
function mapArraysToObjects(data, urlColumn, nameHeader, urlHeader) {
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  const out = [];
  console.time("Total row processing time");
  obj.forEach(function (row, rowIdx) {
    if (row[nameHeader] !== '' && row[urlHeader] === '' && !sheet.isRowHiddenByFilter(rowIdx + 2)) {
      console.time(`Row '${rowIdx + 2}' processing time `);
      try {
        let newFile = template.makeCopy(row[nameHeader].toString(), folder);
        let fileUrl = newFile.getUrl(); // Get the URL of the newly created file
        sheet.getRange(rowIdx + 2, parseInt(urlColumn)).setValue(String(fileUrl)); // Write the URL to the specified urlHeader in the corresponding row
        // formSetup(newFile, responseTarget, newTabName = row[nameHeader].toString());

        out.push([fileUrl]);
        console.info(`Copy created for '${row[nameHeader]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        out.push([e.message] || 'Unknown error occurred');
        console.error(`Failed to create copy for '${row[nameHeader]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      } finally {
        console.timeEnd(`Row '${rowIdx + 2}' processing time `);
      }
    } else {
      if (row[urlHeader] !== '') {
        console.log(`Skipping Row ${rowIdx + 2} - Copy already created.`);
      }
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      out.push([row[urlHeader]] || '');
    }
  });

  return out;
}

// console.time(`START: `); // start a process timer
// console.timeEnd(`END: `); // end a proceess timer
// console.log(`DEBUG: Constant message, followed by variable: '${e.message}'`); // debug
// console.info(`INFO: Constant message, followed by variable: '${e.message}'`); // info
// console.warn(`WARNING: Constant message, followed by variable: '${e.message}'`); // warning
// console.error(`ERROR: Constant message, followed by variable: '${e.message}'`); // error