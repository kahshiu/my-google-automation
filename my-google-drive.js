/*
@OnlyCurrentDoc
*/

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("GDrive Helper")
    .addItem("Set Rootfolder", "captureMyRootfolder")
    .addItem("Set Team", "captureMyGroup")
    .addItem("Create Subfolder", "createMySubfolder")
    .addItem("Create + Share Subfolder + Notification email", "runMyTasks")
    .addToUi();
}

function captureMyRootfolder() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sourceCell = sourceSheet.getActiveCell();
  writeIntoSheet("automation_data", "A1", [
    ["rootfolder", sourceCell.getValue()],
  ]);
}

function captureMyGroup() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sourceCell = sourceSheet.getActiveCell();
  writeIntoSheet("automation_data", "A2", [["group", sourceCell.getValue()]]);
}

function createMySubfolder() {
  const sourceValues = getActiveRangeValues();
  const targetRange = [];

  const [[rootfolderUrl]] = getRangeValues("automation_data", "B1");
  for (let i = 0; i < sourceValues.length; i++) {
    let value = sourceValues[i][0];
    if (i === 0) {
      targetRange.push(["subfolders", value]);
    } else {
      targetRange.push(["", value]);
    }
    createSubfolder(rootfolderUrl, value);
  }
  writeIntoSheet("automation_data", "A3", targetRange);
}

/**
 * tarcAutomation
 * description: create folder, share folder and email relevent students
 */
function runMyTasks() {
  const sourceValues = getActiveRangeValues();
  const studentJson = jsonifyStudentRange(sourceValues);
  const studentEmails = studentJson.map((stu) => stu.studentEmail);
  const [leader] = studentJson.filter(
    (stu) => stu.rank.toLowerCase() === "leader"
  );

  writeIntoSheet("automation_data", "A10", [
    [JSON.stringify(studentJson)],
    [JSON.stringify(studentEmails)],
    [JSON.stringify(leader)],
  ]);

  // NOTE: create group subfolder
  const [[rootfolderUrl]] = getRangeValues("automation_data", "B1");
  const [[groupName]] = getRangeValues("automation_data", "B2");
  const groupFolder = `${groupName}-${leader.studentName}`;
  const { subfolder: subfolder1Obj } = createSubfolder(
    rootfolderUrl,
    groupFolder
  );
  const foldersCreated = [["subfolders", groupFolder]];
  const emailsSent = [];

  // NOTE: emails
  shareFolder(rootfolderUrl, groupFolder, studentEmails);
  const emailSubject = "Shared Folder to Submit Group Assignment";
  const emailBody =
    "Please submit your email assignment here. Individual Components to go into respective student folders, while Shared Components are to stay on base folder. Please visit the folder linke here: " +
    getShareLink(subfolder1Obj);

  studentJson.forEach((stu, key) => {
    // student folder
    const studentFolder = `${groupFolder}/${stu.studentName}`;
    createSubfolder(rootfolderUrl, studentFolder);
    foldersCreated.push(["", studentFolder]);

    // student email
    if (key === 0) {
      emailsSent.push(["emailed", stu.studentEmail]);
    } else {
      emailsSent.push(["", stu.studentEmail]);
    }
  });

  sendMyEmail(studentEmails.join(", "), emailSubject, emailBody);
  writeIntoSheet("automation_data", "A3", foldersCreated);
  writeIntoSheet("automation_data", "C3", emailsSent);
}

/**
 * helper function to convert range into json
 * arguments: studentRange
 * returns: json
 */
function jsonifyStudentRange(studentRange) {
  const headers = ["studentName", "studentEmail", "rank"];
  const values = studentRange;
  const jsonArray = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowObj = {};
    for (var j = 0; j < row.length; j++) {
      var header = headers[j];
      var cellValue = row[j];
      rowObj[header] = cellValue;
    }
    jsonArray.push(rowObj);
  }

  return jsonArray;
}

/**
 * getFolderIdByUrl
 * description: extract folderId from given url
 *    url should be as reflected on folder URL address on browser
 */
function getFolderIdByUrl(targetUrl) {
  // NOTE: URL api is unavailable in app script, as its based on older version of ECMAScript
  const url = targetUrl.split("?")[0];
  const urlObj = {
    protocol: url.split("://")[0],
    host: url.split("/")[2],
    pathname: url.split("/").slice(3),
  };
  const pathname = urlObj.pathname;
  const validPathname = pathname && pathname.length > 0;
  if (!validPathname) return null;

  const folderId = pathname.pop();
  return folderId;
}

/**
 * ensureFolderCreated
 * description: will return folderObject for targeted subfolderName
 *    if folder already exists, will return folderObj of specified name. NOTE: will return first enumerated object only
 *    if folder not exists, will create
 */
function ensureSubfolderCreated(folderObj, subfolderName) {
  const folders = folderObj.getFoldersByName(subfolderName);
  const hasSubfolders = folders.hasNext();
  if (hasSubfolders) return folders.next();
  return folderObj.createFolder(subfolderName);
}

/**
 * ensuredubfolderCreatedByPath
 * description: will return folderObject for targeted subpath, last child node
 *    will ensure all necessary folders are created
 */
function ensureSubfolderCreatedByPath(folderObj, subpath) {
  const isValidSubpath = subpath.length > 0;
  if (!isValidSubpath) return null;

  let parent = folderObj;
  let child = null;
  const subpathSplitted = subpath.split("/");
  while (subpathSplitted.length > 0) {
    const childpath = subpathSplitted.shift();
    child = ensureSubfolderCreated(parent, childpath);
    parent = child;
  }
  return child;
}

/**
 * createSubfolder
 * description: will create subfolders
 * arguments:
 *  url: target url for google drive folder, eg. 'https://drive.google.com/drive/folders/abc123-asdf-abc123-asdf'
 *  folderPath: path to folder, eg. 'group1/student1'
 */
function createSubfolder(url, folderPath) {
  const folderId = getFolderIdByUrl(url);
  if (!folderId) return { message: "Invalid folderId extracted from url" };
  try {
    const folder = DriveApp.getFolderById(folderId);
    const subfolder = ensureSubfolderCreatedByPath(folder, folderPath);
    return { mesage: "success", subfolder };
  } catch (error) {
    return { message: error };
  }
}

/**
 * shareFolder
 * description: share folder with a particular email
 * arguments:
 *  url: target url for google drive folder, eg. 'https://drive.google.com/drive/folders/abc123-asdf-abc123-asdf'
 *  folderPath: path to folder, eg. 'group1/student1'
 *  emails: array of emails to share folder with
 */
function shareFolder(url, folderPath, emails) {
  const folderId = getFolderIdByUrl(url);
  if (!folderId) return { message: "Invalid folderId extracted from url" };

  try {
    const folder = DriveApp.getFolderById(folderId);
    const targetPath = ensureSubfolderCreatedByPath(folder, folderPath);
    Logger.log(targetPath.getName());
    for (const email of emails) {
      targetPath.addEditor(email);
    }

    Logger.log("Folder shared with: " + emails);
  } catch (error) {
    return { message: error };
  }
}

/**
 * sendMyEmail
 * description: send email to targeted client
 * arguments:
 *  recipientEmail: to email
 *  subjectEmail: subject of email
 *  htmlEmail: html of email
 *  options: [optional field]
 *   bccEmail: blind carbon copy to email [optional field]
 */
function sendMyEmail(recipientEmail, subjectEmail, htmlEmail, options) {
  MailApp.sendEmail({
    to: recipientEmail,
    bcc: options?.bccEmail ?? "",
    subject: subjectEmail,
    htmlBody: htmlEmail,
  });

  Logger.log("HTML Email sent to: " + recipientEmail);
}

/**
 * ensure sheet created
 * description: ensure sheet with targeted name is created
 * @param {*} name: string
 * @returns returns sheet object
 */
function ensureSheetCreated(name) {
  const activeSh = SpreadsheetApp.getActiveSpreadsheet();
  let sh = activeSh.getSheetByName(name);
  if (sh) return sh;

  sh = activeSh.insertSheet(name);
  return sh;
}

/**
 * writeIntoSheet
 * description: write an array into a particular sheet's cell
 * @param {*} sheetName: string
 * @param {*} cellName: string, eg. A1, A3, B3
 * @param {*} sourceArray: nested array of rows, eg. [[row1cell1, row1cell2, row1cell3], [row2cell1, row2cell2, row2cell3]]
 */
function writeIntoSheet(sheetName, cellName, sourceArray) {
  const sh = ensureSheetCreated(sheetName);
  const rowCount = sourceArray.length;
  const colCount = sourceArray[0].length;
  const range = sh.getRange(cellName).offset(0, 0, rowCount, colCount);
  range.setValues(sourceArray);
}

// SECTION: UI related
/**
 * getActiveRangeValues
 * description: get active range for further processing
 * @returns rows of values
 */
function getActiveRangeValues() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sourceRange = sourceSheet.getActiveRange();
  const sourceValues = sourceRange.getValues();
  return sourceValues;
}

/**
 * getRangeValues
 * description: get targeted range values
 * @returns rows of values
 */
function getRangeValues(sheetName, notation) {
  const sh = ensureSheetCreated(sheetName);
  return sh.getRange(notation).getValues();
}

/**
 * getSharedLink
 * description: construct google drive shareable link
 * @param {*} folder
 * @returns google drive shareable link
 */
function getShareLink(folder) {
  return `https://drive.google.com/drive/u/1/folders/${folder.getId()}`;
}
