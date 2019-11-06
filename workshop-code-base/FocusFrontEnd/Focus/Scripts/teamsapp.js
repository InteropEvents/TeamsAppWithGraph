(function () {
  'use strict';

  // Call the initialize API first
  microsoftTeams.initialize();

  // Check the initial theme user chose and respect it
  microsoftTeams.getContext(function (context) {
    if (context && context.theme) {
      setTheme(context.theme);
    }
  });

  // Handle theme changes
  microsoftTeams.registerOnThemeChangeHandler(function (theme) {
    setTheme(theme);
  });

  // Save configuration changes
  microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
    // Let the Microsoft Teams platform know what you want to load based on
    // what the user configured on this page
    microsoftTeams.settings.setSettings({
        contentUrl: createTabUrl(), // Mandatory parameter
        entityId: uniqueName() // Mandatory parameter
    });

    // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
    // for the app to call this API before it dismisses the dialog. If the wait times out, you will
    // see an error indicating that the configuration settings could not be saved.
    saveEvent.notifySuccess();
  });

  // Logic to let the user configure what they want to see in the tab being loaded
    document.addEventListener('DOMContentLoaded', function () {

        microsoftTeams.settings.setValidityState(true);

        var tabChoice = document.getElementById('tabChoice');
        if (tabChoice) {
          tabChoice.onchange = function () {
            var selectedTab = this[this.selectedIndex].value;

            // This API tells Microsoft Teams to enable the 'Save' button. Since Microsoft Teams always assumes
            // an initial invalid state, without this call the 'Save' button will never be enabled.
            microsoftTeams.settings.setValidityState(selectedTab === 'first' || selectedTab === 'second');
          };
        }
    });

  // Set the desired theme
  function setTheme(theme) {
    if (theme) {
      // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
      //document.body.className = 'theme-' + (theme === 'default' ? 'light' : theme);
    }
  }

  // Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
  function createTabUrl() {
    var tabChoice = document.getElementById('tabChoice');
    var selectedTab = tabChoice[tabChoice.selectedIndex].value;

    return window.location.protocol + '//' + window.location.host + '/' + "WorkFlow/Index";

    //return window.location.protocol + '//' + window.location.host + '/' + selectedTab;
  }

    function uniqueName() {
      return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
          var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
          return v.toString(16);
      });
  }

})();

$(document).ready(function () {
    microsoftTeams.getContext((context) => {
        teamsContext = context;
        SSO();
    });
});

//https://docs.microsoft.com/en-us/graph/api/channel-post-messages?view=graph-rest-beta&tabs=http
async function sendMessage() {
    btnAnimation();
    var teamId = teamsContext.groupId;
    var channelId = teamsContext.channelId;

    //Please read the graph document to find out the api path
    var path = null;

    var messageContent = $('#Focus_Message_Content').val();

    //Please read the graph document to find out the message body format
    var message = null;
    await graphClientBeta.api(path).post(message);
}

//https://docs.microsoft.com/en-us/graph/api/channel-post-messages?view=graph-rest-beta&tabs=http
async function sendCard() {
    btnAnimation();
    var teamId = teamsContext.groupId;
    var channelId = teamsContext.channelId;

    //Please read the graph document to find out the api path
    var path = null;
    var messageContent = $('#Focus_Message_Content').val();

    //Please read the graph document to find out the message body format
    var message = null;
    await graphClientBeta.api(path).post(message);
}
