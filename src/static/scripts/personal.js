(() => {
  "use strict";

  const params = new URLSearchParams(window.location.search);
  const appId = params.get("appId") || "3d7984b9-9f72-4035-be46-69a152b8f6c2";

  // Call the initialize API first
  microsoftTeams.app.initialize().then(() => {
    microsoftTeams.app.getContext().then((context) => {
      document.getElementById("context").innerHTML = JSON.stringify(context, null, 2);
    });
  });

  const allTabsJson = localStorage.getItem("allTabs");
  const allTabs = allTabsJson ? JSON.parse(allTabsJson) : {};
  Object.keys(allTabs).forEach((pageId) => {
    const channelId = allTabs[pageId].channelId;

    const button = document.createElement("button");
    button.textContent = pageId;
    const h1 = document.querySelector("h1");
    h1.appendChild(button);

    button.addEventListener("click", () => {
      microsoftTeams.pages.navigateToApp({
        appId,
        pageId,
        channelId,
      })
    });
  })

  microsoftTeams.appInitialization.notifySuccess();
})();
