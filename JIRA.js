class JIRA {
  constructor(jiraConfig) {
    if (!jiraConfig instanceof JIRAConfig)
      throw new Error("JIRAConfigクラスを引数に入れてください");
    this.jiraConfig = jiraConfig;
  }

  importPBI() {
    const jql = encodeURI(
      "project = " +
        this.jiraConfig.project() +
        ' AND issuetype in (Story, Task) AND status = "To Do" AND created >= ' +
        this.jiraConfig.startDate() +
        ' AND created <= "' +
        this.jiraConfig.endDate() +
        ' 23:59"'
    );
    const response = UrlFetchApp.fetch(
      "https://" +
        this.jiraConfig.host() +
        "/rest/api/2/search?fields=summary,customfield_10028&jql=" +
        jql,
      {
        contentType: "application/json",
        headers: { Authorization: "Basic " + this.jiraConfig.credential() },
      }
    );

    return JSON.parse(response.getContentText());
  }

  updatePBIStoryPoint(key, point) {
    const response = UrlFetchApp.fetch(
      "https://" +
        this.jiraConfig.host() +
        "/rest/api/3/issue/" +
        key +
        "/editmeta",
      {
        contentType: "application/json",
        headers: { Authorization: "Basic " + this.jiraConfig.credential() },
      }
    );
    const meta = JSON.parse(response.getContentText());
    const pointfield = Object.entries(meta.fields)
      .filter(([_, value]) =>
        ["Story Points", "ストーリーポイント", "Story point estimate"].includes(
          value.name
        )
      )
      ?.reduce((rs, [key, _]) => rs ?? key, null);

    if (!pointfield)
      throw new Error("ストーリーポイントの表示設定をしてください");

    const payload = JSON.stringify({
      notifyUsers: false,
      fields: {
        [pointfield]: point,
      },
    });

    UrlFetchApp.fetch(
      "https://" + this.jiraConfig.host() + "/rest/api/3/issue/" + key,
      {
        contentType: "application/json",
        headers: { Authorization: "Basic " + this.jiraConfig.credential() },
        method: "put",
        payload,
      }
    );
  }
}
