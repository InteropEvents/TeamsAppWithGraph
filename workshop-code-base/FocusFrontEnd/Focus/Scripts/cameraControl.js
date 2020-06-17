// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var audioRecording = false;
var mediaRecorder = null;
var audioChunks = null;
var audioBlob = null;

async function enumerateAndStartCamera() {
    var videoDeviceIndex = 0;
    navigator.mediaDevices.enumerateDevices()
        .then(devices => {
            devices.forEach(function (device) {
                if (device.kind == "videoinput") {
                    console.log(device.kind + ": " + device.label + " id = " + device.deviceId);
                    videoDevices[videoDeviceIndex++] = device.deviceId;
                }
            });

            currentVideoDeviceIndex = 0;
            const hdConstraints = {
                video: { width: { min: 1280, exact: clientWidth }, height: { min: 720, exact: clientHeight } },
                deviceId: { exact: videoDevices[currentVideoDeviceIndex] }
            };

            // Start the Camera
            if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
                navigator.mediaDevices.getUserMedia(hdConstraints).then(function (stream) {
                    video.srcObject = stream;
                    video.play();
                    switchMode('Edit');
                }).catch(function (err) {
                    console.log(err);
                    const updatedConstraints = { video: { width: { min: 1280 }, height: { min: 720 } } };
                    navigator.mediaDevices.getUserMedia(updatedConstraints).then(function (stream) {
                        setClientArea(1280, 720);
                        video.srcObject = stream;
                        video.play();
                        switchMode('Edit');
                    });
                });
            }
        });
}

async function switchCamera() {
    console.log("Switch camera selected");
    currentVideoDeviceIndex = currentVideoDeviceIndex == 0 ? 1 : 0
    const updatedConstraints = {
        video: { deviceId: videoDevices[currentVideoDeviceIndex], width: { exact: clientWidth }, height: { exact: clientHeight } }
    };
    // Start the Camera
    if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
        navigator.mediaDevices.getUserMedia(updatedConstraints).then(function (stream) {
            video.srcObject = stream;
            video.play();
        }).catch(function (err) {
            console.log(err);
            const updatedConstraints = {
                video: { deviceId: videoDevices[currentVideoDeviceIndex], width: { min: 1280 }, height: { min: 720 } }
            };
            navigator.mediaDevices.getUserMedia(updatedConstraints).then(function (stream) {
                video.srcObject = stream;
                video.play();
            });
        });
    }
}

function changeFlipCameraState(bEnable) {
    if (bEnable) {
        $("#Focus_FlipCamera_Img").attr("src", "/Icons/Flipcamera-Enabled.png");
        $('#Focus_FlipCamera').attr("disabled", "true");
        $("#Focus_FlipCamera").attr("disabled", false);

    }
    else {
        $("#Focus_FlipCamera_Img").attr("src", "/Icons/Flipcamera-Disabled.png");
        $("#Focus_FlipCamera").attr("disabled", true);
    }
}

function switchMode(mode) {
    if (mode === states.VIEW) {
        console.log("switchMode: Switching to View Mode");
        StopVideo();
        changeFlipCameraState(false);
        $("#Focus_TakeSnapshot").attr("disabled", true);
        $("#Focus_StartAudio").attr("disabled", true);
        currentState = states.VIEW;
    } else if (mode == states.EDIT) {
        console.log("switchMode: Switching to Edit Mode");
        $("#mainTable").show();
        $("#focusLoader").hide();
        changeFlipCameraState(videoDevices.length > 1);
        $("#Focus_TakeSnapshot_Img").attr("src", "/Icons/Camera-Enabled.png");
        $("#Focus_TakeSnapshot").attr("disabled", false);
        $("#Focus_StartAudio").attr("disabled", true);
        currentState = states.EDIT;
    } else if (mode == states.EDIT_DONOT_ALLOW_CAMERA) {
        console.log("switchMode: Switching to EDIT_DONOT_ALLOW_CAMERA Mode");
        StopVideo();
        changeFlipCameraState(false);
        $("#Focus_TakeSnapshot_Img").attr("src", "/Icons/Camera-Disabled.png");
        $("#Focus_TakeSnapshot").attr("disabled", true);
        $("#Focus_StartAudio").attr("disabled", false);

        // Don't change this, this is intentional
        currentState = states.EDIT;
    }
}

async function initiateAudio() {
    $("#AudioText").hide();
    $("#AudioText").css("top", clientHeight + 5);
    $("#AudioText").css("left", clientWidth * 2.5 / 4);

    if (audioRecording === false) {
        audioRecording = true;
        $("#Focus_StartAudio_Img").attr("src", "/Icons/Mic-Enabled.png");
        navigator.mediaDevices.getUserMedia({ audio: true })
            .then(function (stream) {
                mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/ogg;codecs=opus' }, workerOptions);
                audioChunks = [];

                mediaRecorder.addEventListener("dataavailable", function (event) {
                    audioChunks.push(event.data);
                });

                mediaRecorder.addEventListener("stop", function () {
                    DetectImageObjectByAudio(focusMainImageDataFromCognitive, audioChunks, "audio/ogg; codecs=opus").then(function (result) {
                        if (result.Objects.length > 0)
                            $("#AudioTextLabel").text("Match found for text: " + result.Text);
                        else
                            $("#AudioTextLabel").text("No Match found for text: " + result.Text);

                        $("#AudioText").show();

                        var scalex = clientWidth / focusMainImageDataFromCognitive.metadata.width;
                        var scaley = clientHeight / focusMainImageDataFromCognitive.metadata.height;
                        result.Objects.forEach(function (obj) {
                            console.log("match: " + obj.object);
                            console.log("Original position: " + obj.rectangle.x + "," + obj.rectangle.y + "," + obj.rectangle.h + "," + obj.rectangle.w);
                            console.log("Scaled position: " + obj.rectangle.x * scalex + "," + obj.rectangle.y * scaley + "," + obj.rectangle.h * scaley + "," + obj.rectangle.w * scalex);
                            addRectToCanvas(false, obj.rectangle.x * scalex, obj.rectangle.y * scaley, obj.rectangle.w * scalex, obj.rectangle.h * scaley);
                        });
                    },
                        function (err) {
                            console.log("InitiateAudio: DetectImageObjectByAudio Error -" + err);
                            $("#AudioTextLabel").text("Unable to connect to Cognitive services, Try again !");
                            $("#AudioText").show();
                        })
                });

                mediaRecorder.start();
            });
    }
    else {
        $("#Focus_StartAudio_Img").attr("src", "/Icons/Mic-Disabled.png");
        audioRecording = false;
        //$("#AudioPulse").hide();
        //$("#AudioText").show();
        //$("#AudioTextLabel").text("");
        mediaRecorder.stop();
    }
}