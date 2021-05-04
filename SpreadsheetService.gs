class SpreadsheetService{
  constructor(){
    this.mainReportSsId = "1E92oAoDSuzS_cBoD4D8H2FkNEgiqm9ZBsA9BEYj5TVc"
    this.idSsReportYandex = "1SHnx9GkgMtgzUKQKXLVzXlyTF5JiAViCuYgiekwrO08"
    this.idSsReportGoogle = "1U1HYzgcQEg3Yls7zsJTCD829PFPnnnQRVqd_ueA-140"
    this.idSsReportGoogleAccounts = "1qSwCLsmE7LMf-a4GZKsCBJBLpFnzBlRs7gjMaEDUmSQ"
    this.lengthRowMainReport = 15
    this.matchingColYandex = {
      "Date":{
        colMainReport: "Дата",
        colIndexMainReport:null,
        colIndex:null
      },
      "ClientInfo":{
        colMainReport: "Клиент",
        colIndexMainReport:null,
        colIndex:null
      },
      "CampaignName":{
        colMainReport: "Кампания",
        colIndexMainReport:null,
        colIndex:null
      },
      "CampaignId":{
        colMainReport: "№ Кампании",
        colIndexMainReport:null,
        colIndex:null
      },
      "AdNetworkType":{
        colMainReport: "Тип рекламы",
        colIndexMainReport:null,
        colIndex:null
      },
      "Cost":{
        colMainReport: "Расход",
        colIndexMainReport:null,
        colIndex:null
      },
      "Budget":{
        colMainReport: "Дн. бюджет",
        colIndexMainReport:null,
        colIndex:null
      },
      "Impressions":{
        colMainReport: "Показы",
        colIndexMainReport:null,
        colIndex:null
      },
      "Clicks":{
        colMainReport: "Клики",
        colIndexMainReport:null,
        colIndex:null
      },
      "Conversions":{
        colMainReport: "Конверсии",
        colIndexMainReport:null,
        colIndex:null
      },
      "Amount":{
        colMainReport: "Остаток",
        colIndexMainReport:null,
        colIndex:null
      },
      "Source":{
        colMainReport: "Источник",
        colIndexMainReport:null,
        colIndex:null
      }
    }
    this.matchingColGoogle = {
      "Date":{
        colMainReport: "Дата",
        colIndexMainReport:null,
        colIndex:null
      },
      "CustomerDescriptiveName":{
        colMainReport: "Клиент",
        colIndexMainReport:null,
        colIndex:null
      },
      "CampaignName":{
        colMainReport: "Кампания",
        colIndexMainReport:null,
        colIndex:null
      },
      "CampaignId":{
        colMainReport: "№ Кампании",
        colIndexMainReport:null,
        colIndex:null
      },
      "AdvertisingChannelType":{
        colMainReport: "Тип рекламы",
        colIndexMainReport:null,
        colIndex:null
      },
      "Cost":{
        colMainReport: "Расход",
        colIndexMainReport:null,
        colIndex:null
      },  
      "Amount":{
        colMainReport: "Дн. бюджет",
        colIndexMainReport:null,
        colIndex:null
      },
      "Impressions":{
        colMainReport: "Показы",
        colIndexMainReport:null,
        colIndex:null
      },
      "Clicks":{
        colMainReport: "Клики",
        colIndexMainReport:null,
        colIndex:null
      },
      "Conversions":{
        colMainReport: "Конверсии",
        colIndexMainReport:null,
        colIndex:null
      },
      "VideoQuartile75Rate":{
        colMainReport: "Доп. цели",
        colIndexMainReport:null,
        colIndex:null
      },
      "SpendingLimit":{
        colMainReport: "Остаток",
        colIndexMainReport:null,
        colIndex:null
      },
      "Spent":{
        colMainReport: "Остаток",
        colIndexMainReport:null,
        colIndex:null
      },
      "Source":{
        colMainReport: "Источник",
        colIndexMainReport:null,
        colIndex:null
      },
      "VideoViews":{
        colMainReport: "Клики",
        colIndexMainReport:null,
        colIndex:null
      }
    }
  }
  
  /**
   * Add data accounts to created Spreadsheets
   * @param (object) idsSS - Array ids SpreadSheet created
   * @param (object) objectAccountData - Array objects account data
   * @returns (void)
   */
  addDataReportAccountYandex(idsSS, arrayAccountData){
    for(let idSS of idsSS){
      let ss = SpreadsheetApp.openById(idSS)
      let ssName = ss.getName()
      let sheet = ss.getSheets()[0]
      for(let accountData of arrayAccountData){
        if(ssName !== accountData.clientLogin) continue;
        let data = accountData.content.split("\n");
        for(let row of data){
          let dataRow = row.split("	")
          if(dataRow.length <= 1) continue;
          sheet.appendRow(dataRow)
        }
        
      }
    }
  }

  /**
   * Generate report Yandex. It is sum report every account
   * @param (object) listIdsSs - Array ids spreadsheet every account report
   * @param (object) clientsObject - Object clients data from API Yandex Direct
   * @param (object) companyObject - Object companies data from API Yandex Direct
   * @returns (object) data report without header
  */
  generateReportYandex(listIdsSs, clientsObject, companyObject){
    let ssReportYandex = SpreadsheetApp.openById(this.idSsReportYandex)
    let sheetReportYandex = ssReportYandex.getSheets()[0];
    let arrayOutput = []
    let header = []
    let indexColCompaignId = null

    for(let idSS of listIdsSs){
      let ssReportAccount = SpreadsheetApp.openById(idSS);
      let sheetReportAccount = ssReportAccount.getSheets()[0];

      if(sheetReportYandex.getLastRow() === 0){
        var valuesHeader = sheetReportAccount.getRange(1,1,1,sheetReportAccount.getLastColumn()).getValues().flat()
        valuesHeader = valuesHeader.concat(["ClientId", "ClientInfo", "Amount", "Budget"])
        sheetReportYandex.appendRow(valuesHeader)
      }

      if(header.length === 0){
        var valuesHeader = sheetReportAccount.getRange(1,1,1,sheetReportAccount.getLastColumn()).getValues().flat()
        valuesHeader = valuesHeader.concat(["ClientId", "ClientInfo", "Amount", "Budget"])
        header = valuesHeader
      }

      if(!indexColCompaignId){
        for(let col in header){
          if(header[col] === "CampaignId"){
            indexColCompaignId = col;
            break;
          }
        }
      }

      let lastRow = sheetReportAccount.getLastRow();
      if(lastRow === 1) continue

      let valuesData = sheetReportAccount.getRange(2,1,lastRow - 1, sheetReportAccount.getLastColumn()).getValues();
      for(let row of valuesData){
        for(let col in row){
          if(typeof row[col] === "string"){
            let number = Number(row[col].replace(",", "."))
            if(!isNaN(number)){
              row[col] = number
            }
          }
          
          if(row[col] == "--"){
            row[col] = 0
          }
        }
        let prepareDate = row
        let arrayCustomCols = []
        let companyId = row[indexColCompaignId]
        
        let client = clientsObject[ssReportAccount.getName()]
        arrayCustomCols.push(client.ClientId, client.ClientInfo, client.Amount)

        let getObjectCompany = companyObject[companyId]
        if(getObjectCompany){
          arrayCustomCols.push(getObjectCompany.DailyBudget ? getObjectCompany.DailyBudget.Amount : null)
        }

        prepareDate = prepareDate.concat(arrayCustomCols)
        sheetReportYandex.appendRow(prepareDate)
        arrayOutput.push(prepareDate)
      }
    }
    return arrayOutput
  }

  /**
   * Join report accounts google in one report
   * @returns (object) data report without header
  */
  generateReportGoogle(){
    let ssReportAccounts = SpreadsheetApp.openById(this.idSsReportGoogleAccounts)
    let ssReportGoogle = SpreadsheetApp.openById(this.idSsReportGoogle)
    let listSheets = ssReportAccounts.getSheets();
    let data = []
    for(let sheet of listSheets){
      if(sheet.getName() === "Лист1" || sheet.getLastRow() < 2) continue;
      if(data.length === 0){
        data.push(sheet.getRange(1,1,1,sheet.getLastColumn()).getValues().flat())      
      }
      let values = sheet.getRange(2,1,(sheet.getLastRow()-1), sheet.getLastColumn()).getValues()
      values.forEach(value => data.push(value))
    }
    let firstSheet = ssReportGoogle.getSheets()[0];

    if(firstSheet.getLastRow() === 0){
      firstSheet.getRange(1,1,data.length, data[0].length).setValues(data)
      data = data.slice(1)
    } else {
      data = data.slice(1)
      firstSheet.getRange(firstSheet.getLastRow() + 1,1,data.length, data[0].length).setValues(data)
    }

    return data
  }

  /**
   * Generate report for analytics
   * @param (object) dataYandexReport - data yandex report inserted to table during generation
   * @param (object) dataGoogleReport - data google report inserted to table during generation
   * @returns (void)
   */
  generateMainReport(dataYandexReport, dataGoogleReport){
    let ssYandex = SpreadsheetApp.openById(this.idSsReportYandex)
    let ssGoogle = SpreadsheetApp.openById(this.idSsReportGoogle)
    let ssMainReport = SpreadsheetApp.openById(this.mainReportSsId)
    let sheetMainReport = ssMainReport.getSheets()[0]
    let headerMainReport = sheetMainReport.getRange(1,1,1,sheetMainReport.getLastColumn()).getValues().flat()

    headerMainReport.forEach((value, index) => {

      for(let prop in this.matchingColYandex){
        if(this.matchingColYandex[prop].colMainReport === value){
          this.matchingColYandex[prop].colIndexMainReport = index
          break;
        }
      }

      for(let prop in this.matchingColGoogle){
        if(this.matchingColGoogle[prop].colMainReport === value){
          this.matchingColGoogle[prop].colIndexMainReport = index
          break;
        }
      }

    })

    let arrayToInsert = []

    let sheetYandex = ssYandex.getSheets()[0]
    let sheetGoogle = ssGoogle.getSheets()[0]

    let headerYandexReport = sheetYandex.getRange(1,1,1,sheetYandex.getLastColumn()).getValues().flat();
    headerYandexReport.forEach((value, index) => {
      if(this.matchingColYandex[value]){
        this.matchingColYandex[value].colIndex = index
      }
    })

    let headerGoogleReport = sheetGoogle.getRange(1,1,1,sheetGoogle.getLastColumn()).getValues().flat();
    headerGoogleReport.forEach((value, index) => {
      if(this.matchingColGoogle[value]){
        this.matchingColGoogle[value].colIndex = index
      }
    })

    let lastRowYandexReport = sheetYandex.getLastRow()
    if(lastRowYandexReport > 1){
      for(let row of dataYandexReport){
        let arrayToAdd = new Array(this.lengthRowMainReport)
        for(let colName in this.matchingColYandex){
            if(colName === "Source"){
              let indexMainReport = this.matchingColYandex[colName].colIndexMainReport
              arrayToAdd[indexMainReport] = "Яндекс Директ"
              continue;
            }
            let index = this.matchingColYandex[colName].colIndex
            let indexMainReport = this.matchingColYandex[colName].colIndexMainReport
            arrayToAdd[indexMainReport] = row[index]
        }
        arrayToInsert.push(arrayToAdd)
      }
    }


    let lastRowGoogleReport = sheetGoogle.getLastRow()
    if(lastRowGoogleReport > 1){
      
      for(let row of dataGoogleReport){
        let arrayToAdd = new Array(this.lengthRowMainReport)
        for(let colName in this.matchingColGoogle){

            if(colName === "Clicks"){
                let colIndexTypeCompany = this.matchingColGoogle["AdvertisingChannelType"].colIndex
                let typeCompany = row[colIndexTypeCompany];
                let indexMainReport = this.matchingColGoogle[colName].colIndexMainReport

                if(typeCompany === "Video"){
                  let indexVideoView = this.matchingColGoogle["VideoViews"].colIndex
                  arrayToAdd[indexMainReport] = row[indexVideoView];
                } else {
                  let indexClicks = this.matchingColGoogle[colName].colIndex
                  arrayToAdd[indexMainReport] = row[indexClicks];
                }

                continue
            }

            if(colName === "VideoViews") continue;

            if(colName === "Source"){
              let indexMainReport = this.matchingColGoogle[colName].colIndexMainReport
              arrayToAdd[indexMainReport] = "Google Ads"
              continue;
            }

            if("VideoQuartile75Rate" === colName){
              let videoIndex = this.matchingColGoogle[colName].colIndex
              let clicksIndex = this.matchingColGoogle["VideoViews"].colIndex
              let indexMainReport = this.matchingColGoogle[colName].colIndexMainReport

              let video = Number(row[videoIndex].replace("%", ""))/100
              let clicks = Number(row[clicksIndex])
              arrayToAdd[indexMainReport] = +(video * clicks).toFixed(0)
              continue
            }

            if("SpendingLimit" === colName){
              let spendLimitIndex = this.matchingColGoogle[colName].colIndex
              let spentIndex = this.matchingColGoogle["Spent"].colIndex
              let indexMainReport = this.matchingColGoogle[colName].colIndexMainReport

              let amount = Number(row[spendLimitIndex]) - Number(row[spentIndex])
              arrayToAdd[indexMainReport] = amount
              continue;
            }

            if("Spent" === colName) continue;

            let index = this.matchingColGoogle[colName].colIndex
            let indexMainReport = this.matchingColGoogle[colName].colIndexMainReport
            arrayToAdd[indexMainReport] = row[index]
        }
        arrayToInsert.push(arrayToAdd)
      }
    }

    //sheetMainReport.getRange(2,1,sheetMainReport.getLastRow(), sheetMainReport.getLastColumn()).clearContent()
    sheetMainReport.getRange(sheetMainReport.getLastRow() + 1,1,arrayToInsert.length, this.lengthRowMainReport).setValues(arrayToInsert)
  }
}
