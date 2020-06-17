// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function SSO() {
    let config = {
        clientId: g_appId,
        redirectUri: window.location.origin + "/WorkFlow/SilentEndAuth",       // This should be in the list of redirect uris for the AAD app
        cacheLocation: "localStorage",
        navigateToLoginRequestUrl: false,
    };

    // Setup extra query parameters for ADAL
    // - openid and profile scope adds profile information to the id_token
    // - login_hint provides the expected user name
    if (teamsContext.upn) {
        config.extraQueryParameters = "scope=openid+profile&login_hint=" + encodeURIComponent(teamsContext.upn);
    } else {
        config.extraQueryParameters = "scope=openid+profile";
    }

    let authContext = new AuthenticationContext(config);
    // See if there's a cached user and it matches the expected user
    let user = authContext.getCachedUser();
    console.log("SSO: getCachedUser - ", user);
    if (user) {
        if (user.userName !== teamsContext.upn) {
            console.log("SSO: Clearing Cache");
            // User doesn't match, clear the cache
            authContext.clearCache();
        }
    }

    // Get the id token (which is the access token for resource = clientId)
    let token = authContext.getCachedToken("https://graph.microsoft.com");
    console.log("SSO: getCachedToken - ", token);
    if (token) {
        console.log("SSO: token, checkpoint 1");
        graphAccessToken = token;
        postSigninInit();
    } else {
        console.log("SSO: token, checkpoint 2");
        //SignIn();
        // No token, or token is expired
        authContext._renewToken("https://graph.microsoft.com", function (err, accessToken) {
            if (err) {
                signIn();
                console.log("Renewal failed, show SignIn button: " + err);
            } else {
                graphAccessToken = accessToken;
                postSigninInit();
            }
        });
    }
}

async function signIn() {
    console.log("SignIn button clicked");

    if (graphAccessToken == null) {
        console.log("SignIn Clicked graphAccessToken is null");
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/WorkFlow/InitiateAuth",
            width: 600,
            height: 535,
            successCallback: function (result) {
                console.log("Login succeeded: " + result);
                graphAccessToken = result.accessToken;
                postSigninInit();
            },
            failureCallback: function (reason) {
                console.log("Login failed: " + reason);
            }
        });
    } else {
        console.log("SignIn Clicked graphAccessToken has value");
        //let licenseDetails = await graphClient.api('/subscribedSkus').get();
        //var data="";
        //for (var count = 0; count< licenseDetails.value.length; count ++) {
        //    data += '<li class="ms-ContextualMenu-item"><a class="ms-ContextualMenu-link  tabindex="1">' +
        //        licenseDetails.value[count].skuPartNumber +
        //        '</a><i class="ms-Icon ms-Icon--Accept"></i></li><li class="ms-ContextualMenu-item ms-ContextualMenu-item--divider"></li>';
        //}
        //$('#SubScriptionDetailsMenu').show();
        //$('#SubScriptionDetailsMenu').empty();
        //$('#SubScriptionDetailsMenu').append(data);
    }
}

async function postSigninInit() {
    graphClient = MicrosoftGraph.Client.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: function (authDone) {
            authDone(null, graphAccessToken);
        }
    });

    graphClientBeta = MicrosoftGraph.Client.init({
        defaultVersion: 'beta',
        debugLogging: true,
        authProvider: function (authDone) {
            authDone(null, graphAccessToken);
        }
    });

    userDetails = await graphClient.api("/me").get();

    CreateFocusMetadataFolder();
    updateGroupPeopleSelection();
    populateSendmailContacts();
    retrieveOrCreatePlan();

    //setSignInLabel();
}

// Get the user's profile information from Microsoft Graph
async function setSignInLabel() {
    graphClient = MicrosoftGraph.Client.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: function (authDone) {
            authDone(null, graphAccessToken);
        }
    });
    userDetails = await graphClient.api("/me").get();
    $("#SignInLabel").text(userDetails.displayName);

    // this updates the people I (i.e. "me") work with.
    //updateSocialCircle();
    //updateGroupPeopleSelection();
}

// Get the user's profile information from Microsoft Graph
function getUserProfile(accessToken) {
    $.ajax({
        url: "https://graph.microsoft.com/v1.0/me/",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + accessToken);
        },
        success: function (profile) {
            console.log(profile.displayName);

            // if (profile.userPrincipalName === teamsContext.upn) {
            graphAccessToken = accessToken;
            //setSignInLabel();
            //// Create FocusMetadata Folder
            //CreateFocusMetadataFolder();

            //populateSendmailContacts();
            //updateGroupPeopleSelection();


            //retrieveOrCreatePlan();
            // }
        },
        error: function (xhr, textStatus, errorThrown) {
            console.log("textStatus: " + textStatus + ", errorThrown:" + errorThrown);
        },
    });
}

