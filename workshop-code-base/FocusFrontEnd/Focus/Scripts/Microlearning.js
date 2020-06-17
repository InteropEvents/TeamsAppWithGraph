// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var VideoRepo = {
    "Metadata":
        [
            {
                "tags": "Default",
                "Videos": ["https://web.microsoftstream.com/embed/video/ab24e073-0b6c-46a9-b9d9-f962eb9aeb85?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/e9e25b15-9c2f-4d34-aa8e-31eca96656fa?autoplay=false&amp;showinfo=true"]
            },
            {
                "tags": "Refrigerator",
                "Videos": ["https://web.microsoftstream.com/embed/video/d318df58-507f-450d-b377-3d2983fb0854?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/09ada903-2b26-4b29-9da0-af532c1d17a7?autoplay=false&amp;showinfo=true"]
            },
            {
                "tags": "Microwave",
                "Videos": ["https://web.microsoftstream.com/embed/video/0599ad74-6138-4001-b861-1bea6b644834?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/521bf787-aefb-4026-9322-26b43f112a2c?autoplay=false&amp;showinfo=true"]
            },
            {
                "tags": "Oven",
                "Videos": ["https://web.microsoftstream.com/embed/video/a75fbaed-11fb-4283-9a9e-4fbf2de33928?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/bb963fa1-d267-4233-bbc7-4c086b1e9b0d?autoplay=false&amp;showinfo=true"]
            }
        ]
};

function updateMicrolearningTab(imageDescription) {
    var foundAMatch = false;
    var counter = 0;
    imageDescription = imageDescription.toLowerCase();
    for (counter = 0; counter < VideoRepo.Metadata.length; counter++) {
        if (imageDescription.includes(VideoRepo.Metadata[counter].tags.toLowerCase())) {
            foundAMatch = true;
            break;
        }
    }

    if (foundAMatch) {
        $('#microlearning_video1').attr('src', VideoRepo.Metadata[counter].Videos[0]);
        $('#microlearning_video2').attr('src', VideoRepo.Metadata[counter].Videos[1]);
    } else {
        $('#microlearning_video1').attr('src', VideoRepo.Metadata[0].Videos[0]);
        $('#microlearning_video2').attr('src', VideoRepo.Metadata[0].Videos[1]);
    }

    if (foundAMatch && VideoRepo.Metadata[counter].tags.toLowerCase() !== "Default")
        return VideoRepo.Metadata[counter].tags.toLowerCase();
    else
        return "Unknown";  
}
