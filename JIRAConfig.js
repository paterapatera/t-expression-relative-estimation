class JIRAConfig {
  constructor(sheet) {
    if (!sheet) throw new Error("Sheetクラスを引数に入れてください");
    if (!this.apiKey())
      throw new Error(
        "Apps ScriptでスクリプトプロパティでJIRA_API_KEYを設定してください"
      );
    if (!this.userID())
      throw new Error(
        "Apps ScriptでスクリプトプロパティでJIRA_USER_IDを設定してください"
      );
    if (!this.host())
      throw new Error(
        "Apps ScriptでスクリプトプロパティでJIRA_HOST(例:hoge.atlassian.net)を設定してください"
      );
    if (!this.project())
      throw new Error(
        "Apps ScriptでスクリプトプロパティでJIRA_PROJECTを設定してください"
      );
    this.values = sheet.getDataRange().getValues();

    this.PERIOD_W1 = "1週間";
    this.PERIOD_W2 = "2週間";
    this.PERIOD_CUSTOM = "カスタム";

    if (!this.periodType())
      throw new Error("設定値シートの期間が見つかりませんでした");
  }

  apiKey() {
    return PropertiesService.getScriptProperties().getProperty("JIRA_API_KEY");
  }

  userID() {
    return PropertiesService.getScriptProperties().getProperty("JIRA_USER_ID");
  }

  host() {
    return PropertiesService.getScriptProperties().getProperty("JIRA_HOST");
  }

  credential() {
    return Utilities.base64Encode(this.userID() + ":" + this.apiKey());
  }

  project() {
    return PropertiesService.getScriptProperties().getProperty("JIRA_PROJECT");
  }

  periodType() {
    return this.values[0][0];
  }

  startDate() {
    const date = new Date();

    switch (this.periodType()) {
      case this.PERIOD_W1:
        date.setDate(date.getDate() - 6);
        break;
      case this.PERIOD_W2:
        date.setDate(date.getDate() - 13);
        break;
      default:
        return this.dateFormat(new Date(this.values[0][1]));
    }
    return this.dateFormat(date);
  }

  endDate() {
    const date = new Date();

    switch (this.periodType()) {
      case this.PERIOD_W1:
      case this.PERIOD_W2:
        break;
      default:
        return this.dateFormat(new Date(this.values[0][3]));
    }
    return this.dateFormat(date);
  }

  dateFormat(d) {
    return (
      d.getFullYear() +
      "-" +
      (d.getMonth() + 1).toString().padStart(2, "0") +
      "-" +
      d.getDate().toString().padStart(2, "0")
    );
  }
}
