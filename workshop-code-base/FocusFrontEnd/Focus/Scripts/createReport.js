// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

async function createReport(event) {

    var imageName = event.data.imageName;
    var iRectMD = listOfRectMetadata.findIndex(rmd => { return rmd.imageName === imageName; });
    var smallImageUrl = getSnappedImageFromRect(listOfRectMetadata[iRectMD].rectCoordinates);
    const canvas = document.querySelector('canvas');
    ctx = canvas.getContext("2d");
    var image = ctx.getImageData(0, 0, canvas.width, canvas.height);
    var tempCanvas = document.createElement('canvas');
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    var canvastext = tempCanvas.getContext('2d');
    canvastext.putImageData(image, 0, 0);
    tempCanvas.remove();
    var largeImageUrl = tempCanvas.toDataURL("image/png");
    // Split the base64 string in data and contentType
    var block = smallImageUrl.split(";");
    var blockLg = largeImageUrl.split(";");
    // Get the content type of the image
    var contentType = block[0].split(":")[1]; // In this case "image/gif"
    var contentTypeLg = blockLg[0].split(":")[1]; // In this case "image/gif"
    // Get the real base64 content of the file
    var blob = block[1].split(",")[1];
    var blobLg = blockLg[1].split(",")[1];

    console.log("CreateReport: calling getTasksAndAssignments.");
    // change "true" to "$("sendReportCheckbox").ischecked()
    await getTasksAndAssignments();

    console.log("CreateReport: filling in data.");
    var reportData = {};
    reportData["people"] = gpeopleData;
    reportData["tasks"] = gtaskData;
    reportData["buckets"] = gbucketData;
    reportData["planId"] = focusPlanId;
    reportData["bucketId"] = focusBucketId;
    reportData["planTitle"] = focusPlanTitle;
    reportData["graphToken"] = graphAccessToken;
    reportData["userDetails"] = userDetails;
    // these next two are not correct yet. Need to send the image data probably.
    reportData["snappedImageUrl"] = gsnappedImageWebUrl;
    reportData["imageUrl"] = gimageWebUrl;
    reportData["channelName"] = teamsContext.channelName;
    reportData["groupId"] = teamsContext.groupId;
    reportData["smallPartImageData"] = blob;
    reportData["largePartImageData"] = blobLg;

    console.log("CreateReport: reached generateReport call.");
    $("#menu-mask").show();
    await $.ajax({
        type: 'POST',
        url: window.location.origin + "/report/generateReport",
        data: JSON.stringify(reportData),
        contentType: "application/json",
        dataType: 'json',
        success: function (res) {
            console.log("Called generateReport, got response: " + res);
            $("#menu-mask-wait").hide();
            $("#menu-mask-done").show();
        },
        failure: function (res) {
            console.log("Failed to create report.");
        },
        complete: function () {
            setTimeout(function () {
                $("#menu-mask").hide();
                $("#menu-mask-wait").show();
                $("#menu-mask-done").hide();
            }, 2000);
        }
    });
}