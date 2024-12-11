(() => {
  "use strict";

  const appId = "d7e76c4f-6402-4454-a671-cb443181aad7";

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
