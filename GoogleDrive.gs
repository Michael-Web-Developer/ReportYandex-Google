class GoogleDrive{
  constructor(){
    this.idFolderReportAccount = ID_FOLDER_FILES_REPORT_ACCOUNT
  }

  /**
   * Delete all files from folder report accounts
   * @returns (void)
  */
  deleteFilesReportAccounts(){
    let folder = DriveApp.getFolderById(this.idFolderReportAccount)
    let fileIterator = folder.getFiles()
    while(fileIterator.hasNext()){
      fileIterator.next().setTrashed(true)
    }
  }

  /**
   * Create new file Spreadsheet, report by account
   * @param (object) objectAccounts - Object objects report accounts
   * @returns (object) array ids spreadsheets
  */
  createFileReportAccount(objectAccounts){
    let folder = DriveApp.getFolderById(this.idFolderReportAccount)
    let arrayOutput = []
    
    for(let objectReport in objectAccounts){
      let ss = SpreadsheetApp.create(objectAccounts[objectReport].clientLogin)
      let ss_id = ss.getId()
      let file = DriveApp.getFileById(ss_id)
      file.moveTo(folder)
      /*let result = Drive.Files.insert(
        {
          'title':objectAccounts[objectReport].clientLogin, 
          'mimeType': MimeType.GOOGLE_SHEETS,
        },
        objectAccounts[objectReport].blob,
        {
          convert: true,      
        }
      )

      let fileId = result.id;
      let file = DriveApp.getFileById(fileId)
      file.moveTo(folder)*/
      arrayOutput.push(ss_id)
    }
    
    return arrayOutput
  }
}
