class JIRAConfig {
  constructor(sheet) {
    if (!sheet) throw new Error('Sheetクラスを引数に入れてください')
    this.values = sheet.getDataRange().getValues()

    this.PERIOD_W1 = '1週間'
    this.PERIOD_W2 = '2週間'
    this.PERIOD_CUSTOM = 'カスタム'
  }

  apiKey() {
    return PropertiesService.getScriptProperties().getProperty('JIRA_API_KEY');
  }

  userID() {
    return this.values[5][1]
  }

  host() {
    return this.values[6][1]
  }

  credential() {
    return Utilities.base64Encode(this.userID() + ":" + this.apiKey())
  }

  project() {
    return this.values[7][1]
  }

  periodType() {
    return this.values[0][0]
  }

  startDate() {
    const date = new Date()

    switch (this.periodType()) {
      case this.PERIOD_W1:
        date.setDate(date.getDate() - 6)
        break
      case this.PERIOD_W2:
        date.setDate(date.getDate() - 13)
        break
      default:
        return this.dateFormat(new Date(this.values[0][1]))
    }
    return this.dateFormat(date)
  }

  endDate() {
    const date = new Date()

    switch (this.periodType()) {
      case this.PERIOD_W1:
      case this.PERIOD_W2:
        break
      default:
        return this.dateFormat(new Date(this.values[0][3]))
    }
    return this.dateFormat(date)
  }

  dateFormat(d) {
    return d.getFullYear() + '-'
      + (d.getMonth() + 1).toString().padStart(2, '0') + '-'
      + d.getDate().toString().padStart(2, '0')
  }
}
