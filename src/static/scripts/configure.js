(function () {
  "use strict";

  const onSaveHandler = async (saveEvent) => {
    const host = 'https://live-dingo-upright.ngrok-free.app';
    const navigate = document.getElementById('navigate');
    const entityId = `navigate${navigate.value}`;
    const settings = {
      contentUrl: `${host}/tab`,
      websiteUrl: `${host}/tab`,
      entityId,
      suggestedDisplayName: entityId,
    }
    try {
      await microsoftTeams.pages.config.setConfig(settings);
      microsoftTeams.app.getContext().then((context) => {
        const tallTabsJson = localStorage.getItem('allTabs');
        const tallTabs = tallTabsJson ? JSON.parse(tallTabsJson) : {};
        localStorage.setItem('allTabs', JSON.stringify({...tallTabs, [entityId]: {
          channelId: context.channel.id,
        }}));
      });
      saveEvent.notifySuccess();
    } catch (err) {
      saveEvent.notifyFailure(stringifyError(err));
    };
  };

  microsoftTeams.app.initialize().then(() => {
    microsoftTeams.pages.config.registerOnSaveHandler(onSaveHandler);
    microsoftTeams.pages.config.setValidityState(true)
  });
})();
