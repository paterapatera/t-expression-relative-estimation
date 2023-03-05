class JIRA {
  constructor(jiraConfig) {
    if (!jiraConfig instanceof JIRAConfig) throw new Error('JIRAConfigクラスを引数に入れてください')
    this.jiraConfig = jiraConfig
  }

  importPBI() {
    Logger.log([this.jiraConfig.startDate(), this.jiraConfig.endDate()])
    const jql = encodeURI('project = ' + this.jiraConfig.project() + ' AND issuetype in (Story, Task) AND status = "To Do" AND created >= ' + this.jiraConfig.startDate() + ' AND created <= ' + this.jiraConfig.endDate())
    const response = UrlFetchApp.fetch(
      'https://' + this.jiraConfig.host() + '/rest/api/2/search?fields=summary&jql=' + jql,
      {
        contentType: "application/json",
        headers: { "Authorization": "Basic " + this.jiraConfig.credential() },
      }
    )

    return JSON.parse(response.getContentText())
  }
}
