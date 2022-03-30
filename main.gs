function runReport() {
  try {
    let drive = new GoogleDrive()
    let yandex = new YandexDirect()
    let ssService = new SpreadsheetService()

    let objectAccounts = yandex.getListAccounts()
    objectAccounts = yandex.addAmountClients(objectAccounts)

    let objectCompaigns = yandex.getCompaigns()
    let arrayObjectsAccountsReport = yandex.getReport()

    drive.deleteFilesReportAccounts()

    let ids_ss = drive.createFileReportAccount(arrayObjectsAccountsReport)
    ssService.addDataReportAccountYandex(ids_ss, arrayObjectsAccountsReport)

    let dataWithoutHeaderYandex = ssService.generateReportYandex(ids_ss, objectAccounts, objectCompaigns)
    let dataWithoutHeaderGoogle = ssService.generateReportGoogle()

    ssService.generateMainReport(dataWithoutHeaderYandex, dataWithoutHeaderGoogle)
  } catch(e) {
    Telegram.getService().sendError(CUSTOMER, PROJECT, e, ScriptApp)
  }
}