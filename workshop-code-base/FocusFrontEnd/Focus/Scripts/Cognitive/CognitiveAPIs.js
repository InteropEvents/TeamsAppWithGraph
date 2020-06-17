// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function AnalyzeImage(canvasImgDataUrl) {
    var params = {
        "visualFeatures": "Objects,Categories,Description,Color",
        "language": "en"
    };

    return new Promise((resolve, reject) => {
        $.ajax({
            url: visionUrl + $.param(params),
            headers: { "Content-Type": "application/octet-stream", "Ocp-Apim-Subscription-Key": visionSubScriptionKey },
            type: "POST",
            processData: false,
            data: makeblob(canvasImgDataUrl),
            success: function (data) {
                console.log("AnalyzeImage sucess:", data);
                resolve(data);
            },
            error: function (error) {
                console.log("AnalyzeImage error:", error);
                reject(error);
            }
        });
    });
}

function SpeechToText(audioData, audioType) {
    var audioBlob = new Blob(audioData, { type: audioType });
    var params = {
        "language": "en-US"
    };

    return new Promise((resolve, reject) => {
        $.ajax({
            url: speechToTextUrl + $.param(params),
            headers: { "Content-Type": audioType, "Ocp-Apim-Subscription-Key": speechToTextSubscriptionKey },
            type: "POST",
            processData: false,
            data: audioBlob,
            success: function (data) {
                if (data.RecognitionStatus === "Success") {
                    console.log("Sucess: " + data.DisplayText);
                    resolve(data);
                }
                else {
                    console.log("Error: " + data.RecognitionStatus);
                    reject(data);
                }
            },
            error: function (error) {
                console.log("Error: " + error);
                reject(error);
            }
        });
    });
}

function DetectImageObjectByAudio(imageAnalyzeResult, audioData, audioType) {
    return new Promise((resolve, reject) => {
        SpeechToText(audioData, audioType).then(function (result) {
            var data = [];
            data["Text"] = result.DisplayText;
            data["Objects"] = [];
            var newWord = result.DisplayText.toUpperCase();
            imageAnalyzeResult.objects.forEach(function (elem) {
                if (newWord.indexOf(elem.object.toUpperCase())!=-1) {
                    data.Objects.push(elem);
                }
            });
   
            resolve(data);

        },
            function (err) {
                reject(err);
            });
    });
}

function makeblob(canvasDataUrl) {
    var base64 = ';base64,';

    if (canvasDataUrl.indexOf(base64) === -1) {
        var parts = canvasDataUrl.split(',');
        var contentType = parts[0].split(':')[1];
        var raw = decodeURIComponent(parts[1]);
        return new Blob([raw], { type: contentType });
    }

    var parts = canvasDataUrl.split(base64);
    var contentType = parts[0].split(':')[1];
    var raw = window.atob(parts[1]);
    var rawLength = raw.length;

    var uInt8Array = new Uint8Array(rawLength);

    for (var i = 0; i < rawLength; ++i) {
        uInt8Array[i] = raw.charCodeAt(i);
    }

    return new Blob([uInt8Array], { type: contentType });
}