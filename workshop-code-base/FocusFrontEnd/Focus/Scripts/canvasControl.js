// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

class RectMetadata {
    constructor(imageName, imageDescription, canvasWidth, canvasHeight, planId, taskId, rectCoordinates) {
        this.imageName = imageName;
        this.imageDescription = imageDescription;
        this.orgCanvasWidth = canvasWidth;
        this.orgCanvasHeight = canvasHeight;
        this.planID = planId;
        this.taskId = taskId;
        this.rectCoordinates = rectCoordinates;
    }
}

$("#fileLoader").change(function (e) {
    $("#div_photo").show();
    var reader = new FileReader();
    reader.onload = function (event) {
        var img = new Image();
        img.onload = function () {
            const canvas = document.querySelector('canvas');
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
            //rectsOnCanvas.length = 0;
            canvasForDrawing.clear();
            canvasForDrawing.setBackgroundImage(img.src,
                canvasForDrawing.renderAll.bind(canvasForDrawing),
                {
                    scaleX: canvas.width / img.width,
                    scaleY: canvas.height / img.height
                });

            //rectsOnCanvas.length = 0;
            StopVideo();
        };
        img.src = event.target.result;
    };
    reader.readAsDataURL(e.target.files[0]);

    $("#div_video").hide();
});

async function loadPrimaryImage() {
    var path = "/groups/" + teamsContext.groupId + "/drive/root:/" + teamsContext.channelName + ":/children";

    // Get Focus.jpg file
    try {
        var listResponse2 = await graphClient
            .api(path).filter("startswith(name, 'Focus.jpg')").get();

        console.log("LoadPrimaryImage: Get Focus.jpg - ", listResponse2);

        if (listResponse2.value.length > 0) {
            getImageData(listResponse2.value[0]["@microsoft.graph.downloadUrl"], function (dataUrl) {
                console.log("LoadPrimaryImage: Got file data from OD - ", dataUrl);
                $("#div_photo").show();
                var img = new Image();
                img.onload = function () {
                    const canvas = document.querySelector('canvas');
                    canvas.width = clientWidth;
                    canvas.height = clientHeight;
                    canvasForDrawing.clear();
                    canvasForDrawing.setBackgroundImage(img.src,
                        canvasForDrawing.renderAll.bind(canvasForDrawing),
                        {
                            scaleX: canvas.width / img.width,
                            scaleY: canvas.height / img.height
                        });
                    backgroundPrimaryImage = img;
                    $("#mainTable").show();
                    $("#focusLoader").hide();
                    //$('#idLoadRect').show();
                    loadRects();
                }
                img.src = dataUrl;
                $("#div_video").hide();
                //ChangeFlipCameraState(false);
            });
        }
        else {
            // Enable Camera
            enumerateAndStartCamera();

            console.log("LoadPrimaryImage: No Focus.jpg image exists");
        }
    }
    catch (err) {
        console.log("LoadPrimaryImage: Error - ", err);
    }
}

async function takeSnapshot() {
    const canvas = document.querySelector('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    canvas.getContext('2d').drawImage(video, 0, 0, canvas.width, canvas.height);
    var imgAsDataURL = canvas.toDataURL("image/png");
    var img = new Image();
    img.onload = function () {
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        listOfRectMetadata.length = 0;
        canvasForDrawing.clear();
        canvasForDrawing.setBackgroundImage(img.src,
            canvasForDrawing.renderAll.bind(canvasForDrawing),
            {
                scaleX: canvas.width / img.width,
                scaleY: canvas.height / img.height

            });
        img.setAttribute('crossOrigin', '');
        switchMode("EDIT_DONOT_ALLOW_CAMERA");
        // Uploading Image to Teams
        UploadImage(img.src, "Focus");
    };
    img.src = imgAsDataURL;
}

async function loadRects(isRefresh = false) {
    var path = "/groups/" + teamsContext.groupId + "/drive/root:/" + teamsContext.channelName + "/" + focusMetadataFolderName + ":/children";

    // Restoring initial image
    if (listOfRectMetadata.length > 0 && isRefresh === true) {
        await restoreInitialImage();
        listOfRectMetadata.length = 0;
    }

    // Get all the .json files from "FocusMetadata" folder
    try {
        var listResponse = await graphClient
            .api(path).get();

        console.log("LoadRects: Got all the .json files - ", listResponse);

        if (currentState != states.VIEW && listResponse.value.length > 0) {
            currentState = states.VIEW;
        }

        switchMode(listResponse.value.length > 0 ? "View" : "EDIT_DONOT_ALLOW_CAMERA");
        $("#Focus_AudioButton").show();
        AnalyzeImage(backgroundPrimaryImage.src)
            .then(data => {
                focusMainImageDataFromCognitive = data;
            })
            .catch(err => {
                console.log("oadImage : AnalyzeImage failed for Focus.jpeg file -" + err);
                focusMainImageDataFromCognitive = null;
            })

        var taskArray = await getAllTaskStatii();

        for (var count = 0; count < listResponse.value.length; count++) {
            if (listResponse.value[count].name.includes(".json")) {
                $.ajax({
                    type: 'GET',
                    url: listResponse.value[count]["@microsoft.graph.downloadUrl"],
                    dataType: 'json'
                }).done(function (data) {

                    // Load the rects on Canvas after returning from this function.
                    console.log(
                        "LoadRects: Got file data from OD, pushing it to listOfRectMetadata  - ",
                        data);

                    listOfRectMetadata.push(new RectMetadata(data.imageName, data.imageDescription, data.orgCanvasWidth, data.orgCanvasHeight,
                        data.planID,
                        data.taskId,
                        data.rectCoordinates));

                    // Normalizing the coordinates based on screen size.
                    listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.top =
                        listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.top *
                        (clientHeight / listOfRectMetadata[listOfRectMetadata.length - 1].orgCanvasHeight);

                    listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.height =
                        listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.height *
                        (clientHeight / listOfRectMetadata[listOfRectMetadata.length - 1].orgCanvasHeight);

                    listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.left =
                        listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.left *
                        (clientWidth / listOfRectMetadata[listOfRectMetadata.length - 1].orgCanvasWidth);

                    listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.width =
                        listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.width *
                        (clientWidth / listOfRectMetadata[listOfRectMetadata.length - 1].orgCanvasWidth);

                    // Fill rects grey irrespctive
                    listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.fill = 'rgba(243, 242, 241, 0.5)';

                    // No Task, red border, grey inside
                    // Task in Progress, green border, grey inside
                    // Task Complete , grey border, grey inside
                    if (listOfRectMetadata[listOfRectMetadata.length - 1].taskId != 0 && taskArray.length != 0) {
                        let specificTaskStatus = taskArray.find(x => x.id === listOfRectMetadata[listOfRectMetadata.length - 1].taskId).percentComplete;
                        if (specificTaskStatus === 100) {
                            //listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.fill = 'rgba(243, 242, 241, 0.5)';
                            listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.stroke = "rgba(243, 242, 241)";
                        } else {
                            //listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.fill = 'rgba(124, 252, 0, 0.5)';
                            listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates.stroke = "green";
                        }
                    }
                    var fabricRect = new fabric.Rect(listOfRectMetadata[listOfRectMetadata.length - 1].rectCoordinates);
                    fabricRect.lockMovementX = true;
                    fabricRect.lockMovementY = true;
                    fabricRect.lockUniScaling = true;
                    fabricRect.hasRotatingPoint = false;
                    fabricRect.lockScalingX = true;
                    fabricRect.lockScalingY = true;

                    canvasForDrawing.add(fabricRect);
                    canvasForDrawing.renderAll();
                });
            }
        }
    }
    catch (err) {

    }
}

let inRectIndex = null;

canvasForDrawing.on('mouse:move',
    async function (o) {
        if (!isDown) {
            // Get the current mouse position
            const canvas = document.querySelector('canvas');
            ctx = canvas.getContext("2d");
            var r = canvas.getBoundingClientRect();
            var x = o.e.clientX - r.left;
            var y = o.e.clientY - r.top;
            var showToolTip = false;
            let notInRectCount = 0;

            for (var count = 0; count < listOfRectMetadata.length; count++) {
                if (y > listOfRectMetadata[count].rectCoordinates.top &&
                    y <
                    listOfRectMetadata[count].rectCoordinates.top +
                    listOfRectMetadata[count].rectCoordinates.height + listOfRectMetadata[count].rectCoordinates.strokeWidth &&
                    x > listOfRectMetadata[count].rectCoordinates.left &&
                    x <
                    listOfRectMetadata[count].rectCoordinates.left +
                    listOfRectMetadata[count].rectCoordinates.width + listOfRectMetadata[count].rectCoordinates.strokeWidth) {

                    if (!global_bInsideRect || inRectIndex != count) {
                        // remember currect active rect
                        inRectIndex = count;
                        let timeStamp = new Date().toISOString();
                        $(".ms-PersonaCard").attr("timeStamp", timeStamp);
                        if (listOfRectMetadata[count].taskId != 0) {
                            // Add current task info to dialog for updating
                            listOfRectMetadata[count].timeStamp = timeStamp;
                            populateTaskDialog(listOfRectMetadata[count]);
                            getTaskStatus(count, listOfRectMetadata[count].taskId).then(data => {
                                if (data) {
                                    if (data[0].percentComplete === 100)
                                        ctx.strokeStyle = "rgba(243, 242, 241)";
                                    else
                                        ctx.strokeStyle = "green";
                                    ctx.lineWidth = 8;
                                    ctx.strokeRect(listOfRectMetadata[data[0].rectId].rectCoordinates.left + 3.5,
                                        listOfRectMetadata[data[0].rectId].rectCoordinates.top + 3.5,
                                        listOfRectMetadata[data[0].rectId].rectCoordinates.width,
                                        listOfRectMetadata[data[0].rectId].rectCoordinates.height);
                                }
                            });
                        }
                        else {
                            populateTaskDialog(null);
                            $("#taskDescription").val(listOfRectMetadata[count].imageDescription);
                        }

                        // make sure it only bind once
                        $("#createOrUpdatePlannerTask").off("click");
                        $("#createOrUpdatePlannerTask").on("click",
                            {
                                planId: focusPlanId,
                                imageName: listOfRectMetadata[count].imageName,
                                taskId: listOfRectMetadata[count].taskId,
                                imageDescription: listOfRectMetadata[count].imageDescription
                            },
                            createOrUpdatePlannerTask);

                        $("#createReport").off("click");
                        $("#createReport").on("click",
                            {
                                planId: focusPlanId,
                                imageName: listOfRectMetadata[count].imageName,
                                taskId: listOfRectMetadata[count].taskId,
                                imageDescription: listOfRectMetadata[count].imageDescription
                            },
                            createReport);
                        global_bInsideRect = true;
                    }

                    $(".ms-PersonaCard").show();

                    var normalizedTop = listOfRectMetadata[count].rectCoordinates.top;
                    if (normalizedTop + 456 > clientHeight)
                        normalizedTop = clientHeight - 466;

                    if (normalizedTop < 0) {
                        normalizedTop = 10;
                        if (clientHeight < 466) {
                            $(".ms-PersonaCard").css("bottom", '80px');
                            $(".ms-PersonaCard").css('height', 'auto');
                        }
                    }

                    // Adding additional 43 px to offset canvas padding from outer frame
                    $(".ms-PersonaCard").css("top", normalizedTop + 'px');

                    var leftNum = 0;
                    if (listOfRectMetadata[count].rectCoordinates.left +
                        listOfRectMetadata[count].rectCoordinates.width + 360 > clientWidth) {
                        if (listOfRectMetadata[count].rectCoordinates.left + 360 > clientWidth) {
                            leftNum = listOfRectMetadata[count].rectCoordinates.left - 370;
                        }
                        else {
                            leftNum = listOfRectMetadata[count].rectCoordinates.left + listOfRectMetadata[count].rectCoordinates.width - 370;
                        }
                    }
                    else {
                        leftNum = listOfRectMetadata[count].rectCoordinates.left +
                            listOfRectMetadata[count].rectCoordinates.width + 1;
                    }

                    $(".ms-PersonaCard").css("left",
                        leftNum +
                        'px');

                    var image = getSnappedImageFromRect(listOfRectMetadata[count].rectCoordinates);
                    document.getElementById('snappedImage').setAttribute('src', image);
                    tagFromRect = updateMicrolearningTab(listOfRectMetadata[count].imageDescription);
                    showToolTip = true;

                    break;
                } else {
                    //ctx.lineWidth = 8;
                    //ctx.strokeStyle = "red";
                    //ctx.strokeRect(listOfRectMetadata[count].rectCoordinates.left,
                    //    listOfRectMetadata[count].rectCoordinates.top,
                    //    listOfRectMetadata[count].rectCoordinates.width,
                    //    listOfRectMetadata[count].rectCoordinates.height);

                    notInRectCount++;
                    if (notInRectCount == listOfRectMetadata.length) {
                        global_bInsideRect = false;
                    }
                }
            }

            if (!showToolTip) {
                console.log("not in rectangle");
                //$(".ms-PersonaCard").css("display", "none");
                //$("#CognitiveTextDisplay").hide();
            }
        }
        else {
            if (currentState === states.EDIT && listOfRectMetadata.length < noOfRectsAllowed) {
                var pointer = canvasForDrawing.getPointer(o.e);
                if (origX > pointer.x) {
                    rect.set({ left: Math.abs(pointer.x) });
                }
                if (origY > pointer.y) {
                    rect.set({ top: Math.abs(pointer.y) });
                }
                rect.set({ width: Math.abs(origX - pointer.x) });
                rect.set({ height: Math.abs(origY - pointer.y) });
                canvasForDrawing.renderAll();
            }
        }
    });
let rectArray = [];
canvasForDrawing.on('mouse:up',
    function (o) {
        $("#createOrUpdatePlannerTask").off("click");
        $("#createReport").off("click");
        $(".ms-PersonaCard").css("display", "none");
        if (isDown && currentState === states.EDIT && listOfRectMetadata.length < noOfRectsAllowed && rectsOnCanvas.width > 10) {
            // check if there is rect cross each other
            let hasCross = false;
            rectArray.push(rectsOnCanvas);
            $.each(rectArray, function (a, b) {
                $.each(rectArray, function (c, d) {
                    // corner left top
                    if (b.left > d.left && b.left < d.left + d.width &&
                        b.top > d.top && b.top < d.top + d.height) {
                        hasCross = true;
                        return false;
                    }
                    // corner right top
                    if (b.left + b.width > d.left && b.left + b.width < d.left + d.width &&
                        b.top > d.top && b.top < d.top + d.height) {
                        hasCross = true;
                        return false;
                    }
                    // corner right bottom
                    if (b.left + b.width > d.left && b.left + b.width < d.left + d.width &&
                        b.top + b.height > d.top && b.top < d.top + d.height) {
                        hasCross = true;
                        return false;
                    }
                    // corner left bottom
                    if (b.left > d.left && b.left < d.left + d.width &&
                        b.top + b.height > d.top && b.top < d.top + d.height) {
                        hasCross = true;
                        return false;
                    }
                });
                if (hasCross) {
                    return false;
                }
            });
            if (hasCross) {
                $.each(rectArray, function (a, b) {
                    if (b.left === rectsOnCanvas.left && b.top === rectsOnCanvas.top
                        && b.width === rectsOnCanvas.width && b.height === rectsOnCanvas.height) {
                        rectArray.splice(a, 1);
                    }
                });
                rectsOnCanvas.set({ width: 0 });
                rectsOnCanvas.set({ height: 0 });
                canvasForDrawing.renderAll();
                isDown = false;
                return;
            }

            UploadImage(getSnappedImageFromRect(rectsOnCanvas), "FocusSnappedImage-" + rectsOnCanvas.top.toFixed() + "-" + rectsOnCanvas.left.toFixed(), rectsOnCanvas);
        }
        else if (isDown && currentState === states.EDIT && listOfRectMetadata.length < noOfRectsAllowed && rectsOnCanvas.width <= 10) {
            rectsOnCanvas.set({ width: 0 });
            rectsOnCanvas.set({ height: 0 });
            canvasForDrawing.renderAll();
        }

        isDown = false;
    });

var rect, isDown, origX, origY;
canvasForDrawing.on('mouse:down',
    function (o) {
        isDown = true;
        if (currentState === states.EDIT && listOfRectMetadata.length < noOfRectsAllowed) {
            var pointer = canvasForDrawing.getPointer(o.e);
            origX = pointer.x;
            origY = pointer.y;
            addRectToCanvas(true, pointer.x, pointer.y);
        }
    });

function addRectToCanvas(fromMousedown, x, y, w = 0, h = 0) {
    console.log("AddRectToCanvas: Current state is -", currentState);
    if (fromMousedown) {
        rect = new fabric.Rect({
            left: origX,
            top: origY,
            originX: 'left',
            originY: 'top',
            width: x - origX,
            height: y - origY,
            angle: 0,
            fill: 'rgba(243, 242, 241, 0.5)',
            // opacity: 0.1,
            stroke: 'red',
            strokeWidth: 8,
            hasBorder: true
        });
    }
    else {
        rect = new fabric.Rect({
            left: x,
            top: y,
            originX: 'left',
            originY: 'top',
            width: w,
            height: h,
            angle: 0,
            fill: 'rgba(243, 242, 241, 0.5)',
            //opacity: 0.1,
            stroke: 'red',
            strokeWidth: 8,
            hasBorder: true
        });

    }

    canvasForDrawing.add(rect);
    rectsOnCanvas = rect;

    // Adding restriction to prevent a bug where someone clicks as there is no width of rectangle and we get an exception
    // rectsOnCanvas.width > 10
    if (!fromMousedown && currentState === states.EDIT && listOfRectMetadata.length < noOfRectsAllowed && rectsOnCanvas.width > 10) {
        UploadImage(getSnappedImageFromRect(rectsOnCanvas), "FocusSnappedImage-" + rectsOnCanvas.top.toFixed() + "-" + rectsOnCanvas.left.toFixed(), rectsOnCanvas);
    }
}


function getSnappedImageFromRect(rect) {
    if (rect.width > 0) {
        const canvas = document.querySelector('canvas');
        ctx = canvas.getContext("2d");
        var image = ctx.getImageData(rect.left + 10, rect.top + 10, rect.width - 10, rect.height - 10);
        var tempCanvas = document.createElement('canvas');
        tempCanvas.width = rect.width;
        tempCanvas.height = rect.height;
        var canvastext = tempCanvas.getContext('2d');
        canvastext.putImageData(image, 0, 0);
        tempCanvas.remove();
        return tempCanvas.toDataURL('image/png');
    }
}

//canvasForDrawing.on('click', function (e) {
//    console.log('canvas clicked');
//});