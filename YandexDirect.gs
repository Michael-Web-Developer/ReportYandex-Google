class YandexDirect{
  constructor(){
    this.token = YANDEX_TOKEN
    this.urlApi = "https://api.direct.yandex.com/json/v5"
    this.urlApiv4 = "https://api.direct.yandex.com/live/v4/json/"
  }

  /** 
   * Object objects accounts
   * @returns {object}
  */
  getListAccounts(){
    let endPoint = this.urlApi + "/agencyclients"
    let payload = {
      method: "get",
      params: {
        SelectionCriteria:{
          Archived:"NO"
        },
        FieldNames:["ClientId", "ClientInfo", "Login"]
      }
    }

    let option = {
        method:"POST",
        headers:{
          Authorization: "Bearer " + this.token
        },
        contentType: "application/json",
        payload:JSON.stringify(payload)
    }

    let result = UrlFetchApp.fetch(endPoint, option)

    if(result.getResponseCode() === 200){
        let jsonResult = JSON.parse(result.getContentText())
        
        let clients = jsonResult.result.Clients
        if(clients == null) throw new Error("Не найдено аккантов клиента")
        let outputObject = {}
        for(let client of clients){
          outputObject[client.Login] = client
        }
        return outputObject
    }
    throw new Error("Данные по аккаунтам не получены")
  }

  /** 
   * List objects compaigns
   * @returns {object}
  */
  getCompaigns(){
      let objectAccounts = this.getListAccounts()
      let endPoint = this.urlApi + "/campaigns"
      let payload = {
        method: "get",
        params: {
          SelectionCriteria:{
          },
          FieldNames:[
            "Id", 
            "Name",  
            "State",
            "DailyBudget"
          ]
        }
      }

      let option = {
          method:"POST",
          headers:{
            Authorization: "Bearer " + this.token,
            "Client-Login":null
          },

          contentType: "application/json",
          payload:JSON.stringify(payload)
      }

      let objectOutput = {}
      for(let prop in objectAccounts){
          option.headers["Client-Login"] = objectAccounts[prop].Login
          let response = UrlFetchApp.fetch(endPoint, option)
          let result = JSON.parse(response.getContentText()).result.Campaigns;
          for(let compaign of result){
            if(compaign.State === "ARCHIVED") continue;
            if(compaign.DailyBudget){
              compaign.DailyBudget.Amount = compaign.DailyBudget.Amount/1000000
            }
            objectOutput[compaign.Id] = compaign
          }
      }

      return objectOutput
  }

  /**
   * Add amount to gave object
   * @param (object) objectClients - Object accounts
   * @return (object) Edited objectClients
   */
  addAmountClients(objectClients){
    let payload = {
      "method": "AccountManagement",
      "param": {
          "Action": "Get",
          "SelectionCriteria": {
            "Logins": []
          }
      },
      "locale": "ru",
      "token": "AQAAAAAsXESYAAcMFDxXUtw5-UAqsuCRQdVs24k"
    }

    for (let prop in objectClients){
      payload.param.SelectionCriteria.Logins.push(objectClients[prop].Login)
    }

    let option = {
      method:"POST",
      contentType: "application/json",
      payload:JSON.stringify(payload)
    }

    let response = UrlFetchApp.fetch(this.urlApiv4, option)

    let result = JSON.parse(response.getContentText()).data.Accounts

    for(let account of result){
      objectClients[account.Login].Amount = Number(account.Amount)
    }

    return objectClients
  }
  /** 
   * Array report object every account 
   * @returns {object}
  */
  getReport(){
    let listAccounts = this.getListAccounts()
    let endPoint = this.urlApi + "/reports"
    let payload = {
      method: "get",
      params: {
        SelectionCriteria:{
        },
        FieldNames:[
          'Date',
          'AvgClickPosition',
          'AvgCpc',
          'AvgEffectiveBid',
          'AvgImpressionPosition',
          'AvgPageviews',
          'AvgTrafficVolume',
          'AdNetworkType',
          'BounceRate',
          'Bounces',
          'CampaignId',
          'CampaignName',
          'CampaignType',
          'Clicks',
          'ConversionRate',
          'Conversions',
          'Cost',
          'CostPerConversion',
          'Ctr',
          'GoalsRoi',
          'Impressions',
          'ImpressionShare',
          'Profit',
          'Revenue',
          'Sessions',
          'WeightedCtr',
          'WeightedImpressions'
        ],
        ReportName:"ReportGoogleTable" + Utilities.formatDate(new Date(), "GMT+3", 'dd-MM-yyyy'),
        ReportType:"CAMPAIGN_PERFORMANCE_REPORT",
        DateRangeType:"YESTERDAY",
        Format:"TSV",
        IncludeVAT:"NO"
      }
    }
    /*LAST_7_DAYS*/
    let option = {
        method:"POST",
        headers:{
          "Authorization": "Bearer " + this.token,
          "Accept-Language": "ru",
          "Client-Login":null,
          "processingMode":"online",
          "returnMoneyInMicros": "false",
          "skipReportHeader":"true",
          "skipColumnHeader":"false",
          "skipReportSummary":"true"
        },
        contentType: "application/json",
        payload:JSON.stringify(payload)
    }

    let outputObject = []
    for(let accountObject in listAccounts){
      option.headers["Client-Login"] = listAccounts[accountObject].Login
      let response = UrlFetchApp.fetch(endPoint, option);
      let content = response.getContentText()

      outputObject.push({
        content: content,
        clientLogin: listAccounts[accountObject].Login
      })
    }

    return outputObject
  }
}
