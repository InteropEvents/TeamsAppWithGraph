// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

microsoftTeams.initialize();

// ADAL.js configuration
let config = {
    clientId: g_appId,
    redirectUri: window.location.origin + "/WorkFlow/SilentEndAuth",       // This should be in the list of redirect uris for the AAD app
    cacheLocation: "localStorage",
    navigateToLoginRequestUrl: false
};
let authContext = new AuthenticationContext(config);

if (authContext.isCallback(window.location.hash)) {
    authContext.handleWindowCallback(window.location.hash);
    // Only call notifySuccess or notifyFailure if this page is in the authentication popup
    if (window.opener) {
        if (authContext.getCachedUser()) {
            microsoftTeams.authentication.notifySuccess();
        } else {
            microsoftTeams.authentication.notifyFailure(authContext.getLoginError());
        }
    }
}