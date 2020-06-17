// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function getImageData(url, callback) {
    console.log("getImageData: Getting Image from URL ", url);
    var httpRequest = new XMLHttpRequest();
    httpRequest.onload = function () {
        var fileReader = new FileReader();
        fileReader.onloadend = function () {
            callback(fileReader.result);
        };
        fileReader.readAsDataURL(httpRequest.response);
    };
    httpRequest.open('GET', url);
    httpRequest.responseType = 'blob';
    httpRequest.send();
}

function b64toBlob(b64Data, contentType, sliceSize) {
    contentType = contentType || '';
    sliceSize = sliceSize || 512;
    var byteCharacters = atob(b64Data);
    var byteArrays = [];
    for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
        var slice = byteCharacters.slice(offset, offset + sliceSize);

        var byteNumbers = new Array(slice.length);
        for (var i = 0; i < slice.length; i++) {
            byteNumbers[i] = slice.charCodeAt(i);
        }
        var byteArray = new Uint8Array(byteNumbers);
        byteArrays.push(byteArray);
    }
    var blob = new Blob(byteArrays, { type: contentType });
    return blob;
}
