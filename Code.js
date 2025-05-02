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
 * Extracts the folder ID from a Google Drive folder URL.
 * 
 * This function takes a Google Drive folder URL as input and uses a regular 
 * expression to extract the folder ID. If no valid ID is found, returns "undefined".
 * 
 * @param {string} url - The Google Drive folder URL.
 * @return {string|undefined} The folder ID, or undefined if not found.
 * 
 * @throws {Error} If the input URL is not a valid string.
 * 
 * @example
 * const folderId = getIdFromUrl("https://drive.google.com/drive/folders/1Cf1NbSxGq8po5fMpcwsCOq4Wcj6AwBXt");
 * Logger.log(`Extracted folder ID: ${folderId}`);
 */
function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}(?!.*[-\w]{25,})/).toString();
}

/**
 * Removes the query string from a Google Drive URL.
 * 
 * This function removes any parameters following the '?' character in the URL.
 * 
 * @param {string} url - The Google Drive folder URL.
 * @return {string} The URL without the query string.
 * 
 * @example
 * const cleanUrl = removeQueryFromUrl("https://drive.google.com/drive/folders/1Cf1NbSxGq8po5fMpcwsCOq4Wcj6AwBXt?query=someParam");
 * Logger.log(`Clean URL: ${cleanUrl}`);
 */
function removeQueryFromUrl(url) {
  const myRegex = /\?.*$/;
  if (myRegex.test(String(url))) {
    return String(url).replace(myRegex, '');
  }
  return url;
}

/**
 * Creates copies of a template file with specified names in a destination folder.
 * 
 * Prompts the user for a template file URL, filenames, and a destination folder URL.
 * Creates copies for each filename and saves the URLs back to the specified column in the spreadsheet.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If any of the inputs are invalid or if there is an issue accessing or creating files.
 * 
 * @example
 * copyFiles();
 */
function copyFiles() {
  var templateFileUrl = getUserInput('Enter the template file URL');
  var templateFile = DriveApp.getFileById(getIdFromUrl(templateFileUrl));
  var filenamesSheetUrl = getUserInput('Enter the filenames sheet URL');
  var filenamesTabName = getUserInput('Enter the filenames tab name');
  var filenamesRange = getUserInput('Enter filename list range');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  var urlWriteColumn = getUserInput('Enter column number to write URLs');
  
  var sheet = SpreadsheetApp.openByUrl(filenamesSheetUrl).getSheetByName(filenamesTabName);
  var range = sheet.getRange(filenamesRange);
  var filenames = range.getDisplayValues();
  var destinationFolder = DriveApp.getFolderById(getIdFromUrl(destinationFolderUrl));

  var filteredFilenames = filenames.filter((filename) => filename[0] !== "");

  for (let i = 0; i < filteredFilenames.length; i++) {
    let newFile = templateFile.makeCopy(filteredFilenames[i].toString(), destinationFolder);
    let fileUrl = newFile.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(fileUrl);
  }
}

/**
 * Creates new folders with specified names in a destination folder.
 * 
 * Prompts the user for a Google Sheet URL with folder names and creates new folders
 * in the specified destination folder.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If any of the inputs are invalid or if there are issues creating folders.
 * 
 * @example
 * copyFolders();
 */
function copyFolders() {
  var foldersSheetUrl = getUserInput('Enter the foldernames list sheet URL');
  var foldersTabName = getUserInput('Enter the foldernames sheet tab name');
  var foldernamesRange = getUserInput('Enter filename list range');
  var urlWriteColumn = getUserInput('Enter column number to write URLs');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  
  var sheet = SpreadsheetApp.openByUrl(foldersSheetUrl).getSheetByName(foldersTabName);
  var range = sheet.getRange(foldernamesRange);
  var foldernames = range.getDisplayValues();
  var destinationFolder = DriveApp.getFolderById(getIdFromUrl(destinationFolderUrl));

  var filteredFoldernames = foldernames.filter((foldername) => foldername[0] !== "");

  for (let i = 0; i < filteredFoldernames.length; i++) {
    let newFolder = destinationFolder.createFolder(filteredFoldernames[i].toString());
    let folderUrl = newFolder.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(folderUrl);
  }
}

/**
 * Retrieves the names and URLs of all files within a specified Google Drive folder and writes them to the active sheet.
 * 
 * Prompts the user to enter a public link or ID of a parent folder, retrieves all files,
 * and populates the active Google Sheet with the file names and URLs starting from the second row.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified folder cannot be accessed or if there are issues retrieving files.
 * 
 * @example
 * retrieveFiles();
 */
function retrieveFiles() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getIdFromUrl(parentFolderUrl));
  var files = folder.getFiles();
  var fileData = [];
  
  while (files.hasNext()) {
    var file = files.next();
    var fileUrl = 'https://drive.google.com/open?id=' + file.getId();
    fileData.push([file.getName(), fileUrl]);
  }
  
  if (fileData.length > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(2, 1, fileData.length, fileData[0].length).setValues(fileData);
  }
  
  SpreadsheetApp.getUi().alert('Files retrieved successfully.');
}

/**
 * Retrieves the names and URLs of all subfolders within a specified Google Drive folder and writes them to the active sheet.
 * 
 * Prompts the user to enter a public link or ID of a parent folder, retrieves all subfolders,
 * and populates the active Google Sheet with their names and URLs.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified folder cannot be accessed or if there are issues retrieving subfolders.
 * 
 * @example
 * retrieveFolders();
 */
function retrieveFolders() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getIdFromUrl(parentFolderUrl));
  var subfolders = folder.getFolders();
  var folderData = [];
  
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var folderUrl = 'https://drive.google.com/drive/folders/' + subfolder.getId();
    folderData.push([subfolder.getName(), folderUrl]);
  }
  
  if (folderData.length > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(2, 1, folderData.length, folderData[0].length).setValues(folderData);
  }
  
  SpreadsheetApp.getUi().alert('Folders retrieved successfully.');
}

// console.time(`START: `);
// console.timeEnd(`END: `);
// console.log(`DEBUG: `);
// console.info(`INFO: `);
// console.warn(`WARNING: `);
// console.error(`ERROR: `);