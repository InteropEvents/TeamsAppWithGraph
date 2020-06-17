// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

microsoftTeams.initialize();

// Get the tab context, and use the information to navigate to Azure AD login page
microsoftTeams.getContext(function (context) {
    // ADAL.js configuration
    let config = {
        clientId: g_appId,
        redirectUri: window.location.origin + "/WorkFlow/SilentEndAuth",       // This should be in the list of redirect uris for the AAD app
        cacheLocation: "localStorage",
        navigateToLoginRequestUrl: false,
    };

    // Setup extra query parameters for ADAL
    // - openid and profile scope adds profile information to the id_token
    // - login_hint provides the expected user name
    if (context.upn) {
        config.extraQueryParameters = "scope=openid+profile&login_hint=" + encodeURIComponent(context.upn);
    } else {
        config.extraQueryParameters = "scope=openid+profile";
    }

    // Use a custom displayCall function to add extra query parameters to the url before navigating to it
    config.displayCall = function (urlNavigate) {
        if (urlNavigate) {
            if (config.extraQueryParameters) {
                urlNavigate += "&" + config.extraQueryParameters;
            }
            window.location.replace(urlNavigate);
        }
    }

    // Navigate to the AzureAD login page
    let authContext = new AuthenticationContext(config);
    authContext.login();
});