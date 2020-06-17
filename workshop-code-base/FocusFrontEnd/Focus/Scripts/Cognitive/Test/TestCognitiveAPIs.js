// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

$(document).ready(function () {
    $("#DebugSection").append('<p id="testImageAnalyzeAILable">Upload a image to test the image analyze AI</p>');
    $("#DebugSection").append('<input type="file" id="testImageAnalyzeAIElement"  onchange="TestImageAnalyze()"/>');
    $("#DebugSection").append('<br/><br/>');
    $("#DebugSection").append('<button id="TestSpeechToText" onclick="TestSpeechToText()" style="border:1px solid !important;background-color:#eee !important">Test Speech To Text</button>');
});

var imageAnalyzeData = null;

function TestImageAnalyze() {
    var file = $("#testImageAnalyzeAIElement")[0].files[0];
    var reader = new FileReader();

    reader.addEventListener("loadend", function () {
        AnalyzeImage(reader.result).then(function (result) {
            imageAnalyzeData = result;
            alert("Caption: " + result.description.captions[0].text);
        },
            function (err) {
                alert("Error: " + err.responseText);
            })
    }, false);

    if (file) {
        reader.readAsDataURL(file);
    }

}

if (typeof workerOptions == 'undefined') {
    const workerOptions = {
        OggOpusEncoderWasmPath: 'https://cdn.jsdelivr.net/npm/opus-media-recorder@latest/OggOpusEncoder.wasm',
        WebMOpusEncoderWasmPath: 'https://cdn.jsdelivr.net/npm/opus-media-recorder@latest/WebMOpusEncoder.wasm'
    };
}

window.MediaRecorder = OpusMediaRecorder;
var recording = false;
var mediaRecorder = null;
var audioChunks = null;
var audioBlob = null;

function TestSpeechToText() {
    if (!recording) {
        recording = true;
        document.getElementById("TestSpeechToText").innerText = "Stop";
        navigator.mediaDevices.getUserMedia({ audio: true })
            .then(function (stream) {
                mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/ogg;codecs=opus' }, workerOptions);

                audioChunks = [];

                mediaRecorder.addEventListener("dataavailable", function (event) {
                    audioChunks.push(event.data);
                });

                mediaRecorder.addEventListener("stop", function () {
                    if (imageAnalyzeData === null) {
                        SpeechToText(audioChunks, "audio/ogg; codecs=opus").then(function (result) {
                            alert("Text: " + result.DisplayText);
                        },
                            function (err) {
                                alert("Error: " + err);
                            })
                    }
                    else {
                        DetectImageObjectByAudio(imageAnalyzeData, audioChunks, "audio/ogg; codecs=opus").then(function (result) {
                            alert("Text: " + result.Text);
                            result.Objects.forEach(function (obj) {
                                console.log("match: " + obj.object);
                                console.log("position: " + obj.rectangle.x + "," + obj.rectangle.y + "," + obj.rectangle.h + "," + obj.rectangle.w);
                            });
                        },
                            function (err) {
                                alert("Error: " + err);
                            })
                    }
                });

                mediaRecorder.start();
            });
    } else {
        recording = false;
        document.getElementById("TestSpeechToText").innerText = "Speech To Text";

        mediaRecorder.stop();
    }
}

