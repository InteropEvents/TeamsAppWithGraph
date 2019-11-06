var VideoRepo = {
    "Metadata":
        [
            {
                "tags": "Default",
                "Videos": ["https://web.microsoftstream.com/embed/video/ab24e073-0b6c-46a9-b9d9-f962eb9aeb85?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/e9e25b15-9c2f-4d34-aa8e-31eca96656fa?autoplay=false&amp;showinfo=true"]
            },
            {
                "tags": "Refrigerator",
                "Videos": ["https://web.microsoftstream.com/embed/video/eca401b5-3d82-46e2-9956-9dfe6e78e333?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/1c5e7ee3-7b5f-4e37-b885-338fda628fc0?autoplay=false&amp;showinfo=true"]
            },
            {
                "tags": "Microwave",
                "Videos": ["https://web.microsoftstream.com/embed/video/11f26e9c-3504-453d-9908-7a44c6768d97?autoplay=false&amp;showinfo=true", "https://web.microsoftstream.com/embed/video/e0bbc84d-972c-4863-8700-881986618b5d?autoplay=false&amp;showinfo=true"]
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
