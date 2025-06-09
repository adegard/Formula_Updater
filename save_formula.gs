/**
 * Save/update formula inside gdoc file inside specific folder in Drive
 * 
 * Enable Drive API before
 * (https://stackoverflow.com/questions/62175748/driveapp-error-were-sorry-a-server-error-occurred-please-wait-a-bit-and-try)
 * 
 */

function savemyformula(notetoadd, relatedSSid, text, namefile, folderid) {

  // This is an apps script code to create a google doc inside the same drive folder and paste a text
  // Get the  folder
  var folder = DriveApp.getFolderById(folderid);

  //empty document to start with as template (PLEASE ADD THE FILE ID of a google Document template here below)
  var documentId = DriveApp.getFileById("####DOC_TEMPLATE_ID####").makeCopy(namefile, folder).getId();

  // Add the document to the folder
  //var asFile = DriveApp.getFileById(documentId); // get new doc as a file

  //getting actual body structure
  var body = DocumentApp.openById(documentId).getBody();

  var urlfile = "https://docs.google.com/document/d/" + relatedSSid + "/edit";

  // Paste a text
  body.appendParagraph(namefile);
  body.appendParagraph("*************URL*************");
  body.appendParagraph(urlfile);
  body.appendParagraph("*************FORMULA*************");
  body.appendParagraph(text);
  body.appendParagraph("************COMMENTS**************");
  body.appendParagraph(notetoadd);
}


function testingsaving() {
  savemyformula('This is a sample text that I want to paste in the new document.', 'New Document name');
}


/**
 * Get folder Backup inside same SS folder
 * 
 * https://stackoverflow.com/questions/42180255/find-the-folder-id-of-the-current-google-apps-spreadsheet
 */

function getParentFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  while (folders.hasNext()) {
    return folders.next().getId();
  }
}


/**
 * https://stackoverflow.com/questions/74881094/script-to-identify-if-a-folder-exists-and-if-it-does-not-create-it-in-google-dri
 * 
 */
function createOrGetFolder(folderName, parentFolderId) {
  try {

    var parentFolder = DriveApp.getFolderById(parentFolderId),
      folder;

    if (parentFolder) {
      var video_folder = parentFolder.getFoldersByName(folderName);
      if (video_folder.hasNext()) {

        folder = video_folder.next(); 

      } else {
        folder = parentFolder.createFolder(folderName);

      }
    } else {
      throw new Error("Parent Folder with id: " + parentFolderId + " not found");
    }
    Logger.log("folder: " + folder.getId());
    return folder.getId();
  } catch (error) {
    return error;
  }

}
