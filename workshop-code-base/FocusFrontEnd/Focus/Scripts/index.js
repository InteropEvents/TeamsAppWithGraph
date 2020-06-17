// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const noOfRectsAllowed = 3;
const states = {
    EDIT: 'Edit',
    VIEW: 'View',
    EDIT_DONOT_ALLOW_CAMERA: 'EDIT_DONOT_ALLOW_CAMERA'
};
const focusMetadataFolderName = "FocusMetadata";
var backgroundPrimaryImage = "";

microsoftTeams.initialize();
var graphAccessToken = null;
var graphClient = null;
var graphClientBeta = null; 
var clientWidth = 0;
var clientHeight = 0;
var teamsContext = null;
var canvasForDrawing = new fabric.Canvas('PhotoCanvas', { selection: false });
var rectsOnCanvas;
let listOfRectMetadata = [];
var userDetails = null;
var focusMainImageDataFromCognitive = null;
const workerOptions = {
    OggOpusEncoderWasmPath: 'https://cdn.jsdelivr.net/npm/opus-media-recorder@latest/OggOpusEncoder.wasm',
    WebMOpusEncoderWasmPath: 'https://cdn.jsdelivr.net/npm/opus-media-recorder@latest/WebMOpusEncoder.wasm'
};
window.MediaRecorder = OpusMediaRecorder;
var videoDevices = [0, 0];
var currentVideoDeviceIndex = 0;
var tagFromRect = "Unknown";


// Maintains the currentState of the App. EDIT or VIEW are 2 values it can take.
let currentState = states.EDIT;

var CommandBarElements = document.querySelectorAll(".ms-CommandBar");
for (var i = 0; i < CommandBarElements.length; i++) {
    new fabric['CommandBar'](CommandBarElements[i]);
}

var PersonaCardElement = document.querySelectorAll(".ms-PersonaCard");
for (var i = 0; i < PersonaCardElement.length; i++) {
    new fabric.PersonaCard(PersonaCardElement[i]);
}

var TextFieldElements = document.querySelectorAll(".ms-TextField");
for (var i = 0; i < TextFieldElements.length; i++) {
    new fabric['TextField'](TextFieldElements[i]);
}

//var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
//for (var i = 0; i < DropdownHTMLElements.length; ++i) {
//    var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
//}

var PeoplePickerElements = [].map.call(document.querySelectorAll(".ms-PeoplePicker"), function (element, index) {
    var control = new fabric['PeoplePicker'](element);
    if (index === 0) {
        document.querySelectorAll(".ms-PeoplePicker #taskAssignees")[0].addEventListener("mouseover", function (t) {
            var peopleList = control._peoplePickerMenu.querySelectorAll(".ms-PeoplePicker-result");
            var ids = $('div.ms-PeoplePicker-searchBox').find('.ms-Persona-secondaryText').map(function () {
                return this.innerText;
            });
            ids = Array.prototype.concat.apply([], ids);

            [].forEach.call(peopleList, function (item) {
                item.style.display = "";
            });
            peopleList.forEach(function (item) {
                var val = item.querySelector(".ms-Persona-secondaryText").innerText;
                if (ids.indexOf(val) >= 0) {
                    item.style.display = "none";
                }
            });
        });
    }
    return control;
});


var DatePickerElements = document.querySelectorAll(".ms-DatePicker");
for (var i = 0; i < DatePickerElements.length; i++) {
    new fabric['DatePicker'](DatePickerElements[i]);
}

// Grab elements, create settings, etc.
var video = document.getElementById('video');


function importPhoto() {
    $("#fileLoader").click();
}

function enableCard() {
    $(".ms-PersonaCard-actionDetailBox").show();
}

//https://docs.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
async function sendMail() {
    if (graphAccessToken) {
        btnAnimation();
        var mailSubject = $('#Focus_Mail_Subject').val();
        var mailToAddress = $('.ms-Dropdown-title')[0].innerHTML;
        var mailContent = $('#Focus_Mail_Content').val();

        //Please read the graph document to find out the api path and message body format
        var mailMsg = null;

        var apiPath = null;


        try {
            let response = await graphClient.api(apiPath).post({ message: mailMsg });
            console.log(response);
        } catch (error) {
            throw error;
        }
    } else {
        alert("Please Sign in to your account before Sending an Email");
    }
}

function setClientArea(width = 0, height = 0) {
    if (width !== 0 && height !== 0) {
        clientWidth = width;
        clientHeight = height - 70;
    } else {
        clientWidth = window.innerWidth;
        clientHeight = window.innerHeight - 70;
    }

    console.log("SetClientArea: width, height - ", clientWidth, clientHeight);
}

$(document).ready(function () {
    console.log("ready!");
    setClientArea();

    canvasForDrawing.setHeight(clientHeight);
    canvasForDrawing.setWidth(clientWidth);

    $('#Focus_CameraMenu').hide();
    $('#btnOpenFileDialog').hide();
    //$('#idLoadRect').hide();

    microsoftTeams.getContext((context) => {
        teamsContext = context;
        console.log("Teams Context: ", context);
        console.log("Identifier to know which channel this tab is tied to: ", context.channelId);
        SSO();
    });
});

function populateSendmailContacts() {
    // poplulate the Sendmail tab people dropdown

    // get reference to select element
    var sel = document.getElementById('Focus_SendTo');
    for (let i = sel.options.length - 1; i >= 0; i--) {
        sel.remove(i);
    }

    graphClient
        .api('groups/' + teamsContext.groupId + '/members')
        .get().then(function (people) {
            if ((people !== undefined) && (people !== null)) {

                people.value.forEach(function (person) {
                    // create new option element
                    var opt = document.createElement('option');

                    // create text node to add to option element (opt)
                    opt.appendChild(document.createTextNode(person.userPrincipalName));

                    // set value property of opt
                    opt.value = 'option value';

                    // add opt to end of select box (sel)
                    sel.appendChild(opt);
                });

                var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
                for (var i = 0; i < DropdownHTMLElements.length; ++i) {
                    var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
                }
            }
        });
}

async function CreateFocusMetadataFolder() {
    var path = "/groups/" + teamsContext.groupId + "/drive/root:/" + teamsContext.channelName + ":/children";
    const driveItem = {
        name: focusMetadataFolderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail"
    };

    await graphClient.api(path).post(driveItem).then(res => {
        console.log("CreateFocusMetadataFolder response:", res);
    }).catch(err => {
        console.log("CreateFocusMetadataFolder response:", err);
    });

    await loadPrimaryImage();
}

async function restoreInitialImage() {
    canvasForDrawing.clear();
    canvasForDrawing.setBackgroundImage(backgroundPrimaryImage.src,
        canvasForDrawing.renderAll.bind(canvasForDrawing),
        {
            scaleX: clientWidth / backgroundPrimaryImage.width,
            scaleY: clientHeight / backgroundPrimaryImage.height
        });
}

async function UploadImage(imageDataUrl, fileName, rectMetadata = null) {
    // Split the base64 string in data and contentType
    var block = imageDataUrl.split(";");
    // Get the content type of the image
    var contentType = block[0].split(":")[1]; // In this case "image/gif"
    // Get the real base64 content of the file
    var blob = b64toBlob(block[1].split(",")[1], contentType);

    // Create "Focus" Image file in root folder
    if (rectMetadata === null) {
        try {
            let createResponse = await graphClient
                .api("/groups/" +
                    teamsContext.groupId +
                    "/drive/root:/" +
                    teamsContext.channelName +
                    "/" +
                    fileName +
                    ".jpg" +
                    ":/content").put(blob);
            console.log("UploadImage create file response: ", createResponse);
            if (filename.includes("Snap"))
                gsnappedImageWebUrl = createResponse.webUrl;
            else
                gimageWebUrl = createResponse.webUrl;

        } catch (error1) {
            console.log("UploadImage create file error: ", error1);
        }

        AnalyzeImage(imageDataUrl)
            .then(data => {
                focusMainImageDataFromCognitive = data;
            })
            .catch(err => {
                console.log("UploadImage : AnalyzeImage failed for Focus.jpeg file -" + err);
                focusMainImageDataFromCognitive = null;
            })
    }
    else {

        try {
            let createResponse = await graphClient
                .api("/groups/" +
                    teamsContext.groupId +
                    "/drive/root:/" +
                    teamsContext.channelName +
                    "/" +
                    focusMetadataFolderName +
                    "/" +
                    fileName +
                    ".jpg" +
                    ":/content").put(blob);
            console.log("UploadImage create file response: ", createResponse);
        } catch (error1) {
            console.log("UploadImage create file error: ", error1);
        }

        // Create Metadata file (.json format)
        AnalyzeImage(imageDataUrl)
            .then(data => {
                var imageDesc = "";
                // Got the text from Cognitive services
                if (data.description.captions.length > 0)
                    imageDesc = data.description.captions[0].text;
                else
                    imageDesc = "AnalyzeImage API succedded but unable to get text for this image";

                // Got the text from Cognitive services
                var metaDataObj = new RectMetadata(fileName + ".jpeg", imageDesc, clientWidth, clientHeight, "0", "0", rectMetadata);

                listOfRectMetadata.push(metaDataObj);
                try {
                    let createResponse = graphClient
                        .api("/groups/" +
                            teamsContext.groupId +
                            "/drive/root:/" +
                            teamsContext.channelName +
                            "/" +
                            focusMetadataFolderName +
                            "/" +
                            fileName +
                            ".json" +
                            ":/content").put(JSON.stringify(metaDataObj));
                    console.log("UploadImage create file response: ", createResponse);
                } catch (error2) {
                    console.log("UploadImage create file error: ", error2);
                }
            })
            .catch(error => {
                // Unable to get text from Cognitive services
                var metaDataObj = new RectMetadata(fileName + ".jpeg", "Either visionSubScriptionKey value is null in Credentials.js OR The image is not recognized by cognitive services ", clientWidth, clientHeight, "0", "0", rectMetadata);
                listOfRectMetadata.push(metaDataObj);
                try {
                    let createResponse = graphClient
                        .api("/groups/" +
                            teamsContext.groupId +
                            "/drive/root:/" +
                            teamsContext.channelName +
                            "/" +
                            focusMetadataFolderName +
                            "/" +
                            fileName +
                            ".json" +
                            ":/content").put(JSON.stringify(metaDataObj));
                    console.log("UploadImage create file response: ", createResponse);
                } catch (error2) {
                    console.log("UploadImage create file error: ", error2);
                }
            })
    }
}

function StopVideo() {
    console.log("Stopping Video");
    if (video.srcObject) {
        video.srcObject.getTracks().forEach(function (track) {
            track.stop();
        });
    }
    $("#div_video").hide();
    $("#div_photo").show();
}

async function btnAnimation() {
    $("#menu-mask").show();

    setTimeout(function () {
        $("#menu-mask-wait").hide();
        $("#menu-mask-done").show();
    }, 2000);

    setTimeout(function () {
        $("#menu-mask").hide();
        $("#menu-mask-wait").show();
        $("#menu-mask-done").hide();
    }, 3000);
}