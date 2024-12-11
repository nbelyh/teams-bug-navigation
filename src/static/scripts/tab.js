(function () {
  "use strict";

  // Call the initialize API first
  microsoftTeams.app.initialize().then(function () {
    microsoftTeams.app.getContext().then(function (context) {
      document.getElementById("context").innerHTML = JSON.stringify(context, null, 2);
    });
  });

  microsoftTeams.appInitialization.notifySuccess();
})();
