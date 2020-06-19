// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function checkInventory() {
    console.log("check inventory clicked");
    $.ajax({
        type: 'POST',
        //data: JSON.stringify({ "TeamId": teamsContext.groupId, "ChannelId": teamsContext.channelId, "Tag": tagFromRect}),
        contentType: "application/json",
        dataType: 'json',
        url: 'https://prod-17.westus.logic.azure.com:443/workflows/6c7bae7958fe494f8b38bad39596084c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fAj8OKMQVg2aSakMuoidbc7mQH5hxEwMWchbPgiYfzI',
        success: function (data) {
            console.log("checkInventory : Flow success");
        }
    });
}

function startApproval() {
    console.log("startApproval clicked");
    $.ajax({
        type: 'POST',
        contentType: "application/json",
        dataType: 'json',
        data: JSON.stringify({ "TeamId": teamsContext.groupId, "ChannelId": teamsContext.channelId, "Tag": tagFromRect }),
        url: 'https://prod-115.westus.logic.azure.com:443/workflows/75642955ea3544f98da60f008121d5a5/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jtW7FfLkZbaXs5d-BDDytsyPo4VJ2Ab71uikEFWGBlo',
        success: function (data) {
            console.log("startApproval : Flow success");
        }
    });
}

function TriggerFlows() {
    console.log("Tigger flows");
    btnAnimation();
    if ($("#ck_flow_inventory").prop("checked") === true) {
        checkInventory();
    }

    if ($("#ck_flow_approval").prop("checked") === true) {
        startApproval();
    }
}