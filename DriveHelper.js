/**
 * @OnlyCurrentDoc
 * Indicates that this automation will only attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the limited scope.
 */
  
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
  
  // console.time(`START: `); // start a process timer
  // console.timeEnd(`END: `); // end a proceess timer
  // console.log(`DEBUG: Constant message, followed by variable: '${e.message}'`); // debug
  // console.info(`INFO: Constant message, followed by variable: '${e.message}'`); // info
  // console.warn(`WARNING: Constant message, followed by variable: '${e.message}'`); // warning
  // console.error(`ERROR: Constant message, followed by variable: '${e.message}'`); // error