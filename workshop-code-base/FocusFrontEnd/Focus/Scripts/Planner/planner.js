var focusPlanTitle = "Focus Plan";
var focusPlanId = null;
var focusBucketId = null;
var global_bInsideRect = false;
var gtaskData = {};
var gpeopleData = {};
var gGroupPeople = {};
var gbucketData = null;
var gsnappedImageWebUrl = null;
var gimageWebUrl = null;

async function populateTaskDialog(rectmd) {
    $('#createOrUpdatePlannerTask').css("display", "none");
    if (gbucketData === null) {
        var buckets = await graphClient.api("planner/plans/" + focusPlanId + "/buckets").get();
            if (buckets === undefined) {
                console.log("populateTaskDialog(): bucket information not found", task);
                return;
            }
        gbucketData = buckets.value;
    }
    if (rectmd === null) {
        // set defaults for task dialog.
        $("#taskTitle").val("");
        $("#newDueDate").val(new Date().toDateString());
        $("#taskDescription").val("");
        $("#taskAssignees").val("");
        $('div.ms-PeoplePicker-searchBox').find('.ms-Persona--xs').each(function () {
            this.remove();
        });
        updateGroupPeopleSelection();
        $('#createOrUpdateTaskBtn').text("Create Task");
    }
    else {
        if (rectmd.taskId === 0) {
            console.log("populateTaskDialog(): no existing task provided", rectmd);
            $('#createOrUpdateTaskBtn').text("Create Task");
        }
        else {
            // retrieve task and populate the task dialog
            var task = await graphClient.api("planner/tasks/" + rectmd.taskId).get();
            if (task === undefined) {
                console.log("populateTaskDialog(): task not found", task);
                return;
            }
            else {
                var taskDetails = await graphClient.api("planner/tasks/" + rectmd.taskId + "/details").get();
                if (taskDetails === undefined) {
                    console.log("populateTaskDialog(): taskDetails not found", taskDetails);
                    return;
                }
            }
            $('#createOrUpdateTaskBtn').text("Update Task");
        }

        $("#taskTitle").val(task.title);
        var newDueDate = new Date(task.dueDateTime);
        $("#newDueDate").val(newDueDate.toDateString());
        $("#taskDescription").val(taskDetails.description);

        // list of assigned users for this task.
        var userGuids = Object.keys(task.assignments);

        $('div.ms-PeoplePicker-searchBox').find('.ms-Persona-secondaryText').each(function () {
            var i = userGuids.findIndex(g => { return this.innerText === g });
            if (i > -1)
                delete userGuids[i];
        });

        // add the users to the searchbox.
        userGuids.forEach(async function (userGuid) {

            var user = await graphClient.api("users/" + userGuid).get();
            if (user === undefined) {
                console.log("populateTaskDialog(): assinged user not found. ", user);
            }
            else {
                $('div.ms-PeoplePicker-searchBox').prepend(
                    '<div id="assignee-' + userGuid + '" class="ms-Persona ms-Persona--xs ms-Persona--token ms-PeoplePicker-persona">\
                    <div class="ms-Persona-imageArea">\
                        <div class="ms-Persona-initials ms-Persona-initials--blue">TJ</div>\
                    </div>\
                    <div class="ms-Persona-presence">\
                    </div>\
                    <div class="ms-Persona-details">\
                        <div class="ms-Persona-primaryText">' + user.displayName + '</div>\
                        <div class="ms-Persona-secondaryText">' + userGuid + '</div>\
                    </div >\
                    <div class="ms-Persona-actionIcon">\
                        <i id="'+ userGuid + '" class= "ms-Icon ms-Icon--Cancel" ></i >\
                    </div >\
                </div > ');
                $('#' + userGuid).on("click", { id: userGuid }, function (event) {
                    $('#assignee-' + event.data.id).remove();
                });

            }
        });

    }

    $('#createOrUpdatePlannerTask').css("display", "inline");
}

function createOrUpdatePlannerTask(event) {
    if (graphAccessToken) {

        var planId = event.data.planId;
        var imageName = event.data.imageName;
        var assignedUser = null;
        var dueDate = getDateTimeOffset($('#newDueDate').val());

        var iRectMD = listOfRectMetadata.findIndex(rmd => { return rmd.imageName === imageName; });

        var plannerTask = {
            planId: planId,
            bucketId: focusBucketId,
            title: $('#taskTitle').val(),
            dueDateTime: dueDate,
            assignments: {
                //"4e98f8f1-bb03-4015-b8e0-19bb370949d8": {
                //    "@odata.type": "microsoft.graph.plannerAssignment",
                //    "orderHint": "String"
            }
        }

        var plannerTaskForUpdate = {
            title: $('#taskTitle').val(),
            dueDateTime: dueDate,
            assignments: {
                //"4e98f8f1-bb03-4015-b8e0-19bb370949d8": {
                //    "@odata.type": "microsoft.graph.plannerAssignment",
                //    "orderHint": "String"
            }
        }

        var numAssignees = 0;

        // Add assignees from the People Picker search box.
        $('div.ms-PeoplePicker-searchBox').find('.ms-Persona-secondaryText').each(function () {
            numAssignees++;
            plannerTask.assignments[this.innerText] =
                {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !"
                }
        });

        // if no assignees in the search box, just assign to me.
        if (numAssignees == 0) {
            plannerTask.assignments[userDetails.id] =
                {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !"
                }
        }

        // this is for a new task. note the ".post()" and "/planner/tasks"
        if (event.data.taskId === "0") {
            graphClient.api('/planner/tasks').post(plannerTask).then(async function (task) {
                console.log("result creations of task: " + task);
                var det = await graphClient
                    .api('/planner/tasks/' + task.id + '/details')
                    .get();
                let newDescription = $('#taskDescription').val();
                if (newDescription === null)
                    newDescription = event.data.imageDescription;

                // it just happens that the details object id is the task id.
                //var bindid = addLocation($('#newTitle').val(), det.id);

                var reference = {
                    "previewType": "noPreview",
                    "description": newDescription,
                    "references": {
                    }
                };
                //reference.references[encodedDocUrl] =
                //    {
                //        "@odata.type": "microsoft.graph.plannerExternalReference",
                //        "alias": bindid,
                //        "previewPriority": ' !',
                //        "type": Office.context.host
                //    };

                await graphClient
                    .api('/planner/tasks/' + task.id + '/details')
                    .header("If-Match", det["@odata.etag"])
                    .header("Content-Type", "application/json")
                    .patch(reference);

                // now update the json file in OneDrive with the RectMetaData information.
                //var rectMD = listOfRectMetadata.find(rmd => { return rmd.imageName === imageName; });
                listOfRectMetadata[iRectMD].imageDescription = newDescription;
                if (listOfRectMetadata[iRectMD] === undefined)
                    console.log("createOrUpdatePlannerTask(): couldn't find RectMetaData for this task!");
                else {
                    listOfRectMetadata[iRectMD].taskId = task.id;
                    try {
                        var createResponse = await graphClient
                            .api("/groups/" +
                                teamsContext.groupId +
                                "/drive/root:/" +
                                teamsContext.channelName +
                                "/" +
                                focusMetadataFolderName +
                                "/" +
                                listOfRectMetadata[iRectMD].imageName.split(".")[0] +
                                ".json" +
                                ":/content").put(JSON.stringify(listOfRectMetadata[iRectMD]));
                        console.log("UploadImage create file response: ", createResponse);
                    } catch (error) {
                        console.log("UploadImage create file error: ", error);
                    }

                    $('#createOrUpdateTaskBtn').val("Update Task");
                }
            });
        }
        // this is for an existing task. note the ".update()" and "/planner/tasks/{id}"
        else {
            graphClient.api('/planner/tasks/' + event.data.taskId).get().then(async function (task) {
                //var newtask = await graphClient.api('/planner/tasks/' + event.data.taskId).header("If-Match", task["@odata.etag"]).update(plannerTask);

                var numNewAssignees = 0; // reinitialize.
                // Add assignees from the People Picker search box.
                $('div.ms-PeoplePicker-searchBox').find('.ms-Persona-secondaryText').each(function () {
                    numNewAssignees++;
                    plannerTaskForUpdate.assignments[this.innerText] =
                        {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !"
                        }
                });

                // if no assignees in the search box, just assign to me.
                if (numNewAssignees == 0) {
                    plannerTaskForUpdate.assignments[userDetails.id] =
                        {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !"
                        }
                }

                // because this is an existing task, we need to trim out any existing assignees 
                // who are not in our new assignment list.

                var assignees = Object.keys(task.assignments);
                assignees.forEach(aguid => {
                    if (plannerTaskForUpdate.assignments[aguid] == undefined) {
                        plannerTaskForUpdate.assignments[aguid] = null;
                    }
                });


                // this time, fill in the new assignments.
                await graphClient
                    .api('/planner/tasks/' + event.data.taskId)
                    .header("If-Match", task["@odata.etag"])
                    .header("Content-Type", "application/json")
                    .patch(plannerTaskForUpdate);

                var det = await graphClient
                    .api('/planner/tasks/' + task.id + '/details')
                    .get();

                let newDescription = $('#taskDescription').val();

                listOfRectMetadata[iRectMD].imageDescription = newDescription;

                var reference = {
                    "previewType": "noPreview",
                    "description": newDescription,
                    "references": {
                    }
                };
                //reference.references[encodedDocUrl] =
                //    {
                //        "@odata.type": "microsoft.graph.plannerExternalReference",
                //        "alias": bindid,
                //        "previewPriority": ' !',
                //        "type": Office.context.host
                //    };

                await graphClient
                    .api('/planner/tasks/' + task.id + '/details')
                    .header("If-Match", det["@odata.etag"])
                    .header("Content-Type", "application/json")
                    .patch(reference);

                // now update the json file in OneDrive with the RectMetaData information.
                var rectMD = listOfRectMetadata.find(rmd => { return rmd.imageName === imageName });
                if (rectMD === undefined)
                    console.log("createOrUpdatePlannerTask(): couldn't find RectMetaData for this task!");
                else {
                    rectMD.taskId = task.id;
                    try {
                        var createResponse = await graphClient
                            .api("/groups/" +
                                teamsContext.groupId +
                                "/drive/root:/" +
                                teamsContext.channelName +
                                "/" +
                                focusMetadataFolderName +
                                "/" +
                                rectMD.imageName.split(".")[0] +
                                ".json" +
                                ":/content").put(JSON.stringify(rectMD));
                        console.log("UploadImage create file response: ", createResponse);
                    } catch (error) {
                        console.log("UploadImage create file error: ", error);
                    }
                }
            });
        }

        // remove the event handler for now. we'll turn it on when we enter the rect again.
        $("#createOrUpdatePlannerTask").off("click");
    }
}

async function updateSocialCircle() {
    if (graphAccessToken) {
        // use Graph's "social intelligence" to get the people I work with.
        graphClient
            .api('/me/people')
            .get().then(async function (res) {

                console.log(res); // print out the people
                myPeople = res;

                //updatePeoplePickerResultGroup();
                var groupTitleDiv = '<div class="ms-PeoplePicker-resultGroupTitle"> Contacts </div >';
                $('#myPeopleResultGroup').empty();
                $('#myPeopleResultGroup').append(groupTitleDiv);

                var divPeopleResults = myPeople.value.map(function (person) {
                    // TODO: chose this icon based on creator's relationship to me.
                    if ((person.personType.class == "Person") && (person.personType.subclass == "OrganizationUser")) {

                        //var email = person.scoredEmailAddresses[0].address;
                        var fullName = person.displayName;
                        var firstInitial = (person.givenName != null) ? person.givenName.substring(0, 1) : "?";
                        var secondInitial = (person.surname != null) ? person.surname.substring(0, 1) : "?";
                        var initials = firstInitial + secondInitial;
                        //var departmentOrJob = (person.department != null) ? person.department : ((person.jobTitle != null) ? person.jobTitle : "unknown");

                        return '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + initials + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + fullName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + person.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';
                    }
                    else
                        return null;
                });

                divPeopleResults.forEach(function (per) {
                    if (per != null)
                        $('#myPeopleResultGroup').append(per);
                });

                var myFirstInitial = (userDetails.givenName != null) ? userDetails.givenName.substring(0, 1) : "?";
                var mySecondInitial = (userDetails.surname != null) ? userDetails.surname.substring(0, 1) : "?";
                var meDiv = '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + myFirstInitial + mySecondInitial + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + userDetails.displayName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + userDetails.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';

            });
    }
    else
        console.log("no graph token available, something went wrong!", graphAccessToken);
}

async function updateGroupPeopleSelection() {
    if (graphAccessToken) {
        graphClient
            .api('groups/' + teamsContext.groupId + '/members')
            .get().then(function (res) {
                if ((res !== undefined) && (res !== null)) {

                    console.log(res); // print out the people
                    myPeople = res;
                    gGroupPeople = myPeople;

                    //updatePeoplePickerResultGroup();
                    var groupTitleDiv = '<div class="ms-PeoplePicker-resultGroupTitle"> Contacts </div >';
                    $('#myPeopleResultGroup').empty();
                    $('#myPeopleResultGroup').append(groupTitleDiv);

                    var divPeopleResults = myPeople.value.map(function (person) {
                        // TODO: chose this icon based on creator's relationship to me.
                        //if ((person.personType.class == "Person") && (person.personType.subclass == "OrganizationUser")) {

                        //var email = person.scoredEmailAddresses[0].address;
                        var fullName = person.displayName;
                        var firstInitial = (person.givenName != null) ? person.givenName.substring(0, 1) : "?";
                        var secondInitial = (person.surname != null) ? person.surname.substring(0, 1) : "?";
                        var initials = firstInitial + secondInitial;
                        var departmentOrJob = (person.department != null) ? person.department : ((person.jobTitle != null) ? person.jobTitle : "unknown");

                        return '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + initials + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + fullName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + person.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';
                        //}
                        //else
                        //    return null;
                    });

                    divPeopleResults.forEach(function (per) {
                        if (per != null)
                            $('#myPeopleResultGroup').append(per);
                    });

                    var myFirstInitial = (userDetails.givenName != null) ? userDetails.givenName.substring(0, 1) : "?";
                    var mySecondInitial = (userDetails.surname != null) ? userDetails.surname.substring(0, 1) : "?";
                    var meDiv = '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + myFirstInitial + mySecondInitial + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + userDetails.displayName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + userDetails.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';

                    $('#myPeopleResultGroup').append(meDiv);
                }
            });
    }
    else
        console.log("no graph token available, something went wrong!", graphAccessToken);
}

async function retrieveOrCreatePlan() {
    // Look for existing plan. If not found, create it.
    var res = await graphClient.api('groups/' + teamsContext.groupId + '/planner/plans')
        .get();

    var groupres = await graphClient.api('groups/' + teamsContext.groupId)
        .get();

    if (groupres === undefined) {
        console.log("retrieveOrCreatePlan(): error getting group name!");
        return;
    }
    else
        focusPlanTitle = groupres.displayName + " Plan";

    var plan = res.value.find(p => { return p.title === groupres.displayName + " Plan"; });

    if ((res.value.length >= 1) && (plan !== undefined)) {
        focusPlanId = plan.id;
        console.log("SignIn(): Plan found! Plan Id: " + plan.id + " Plan Title: " + plan.title);
    }
    else {

        var plannerBody = {
            "owner": teamsContext.groupId,
            "title": focusPlanTitle
        };

        plan = await graphClient.api("planner/plans")
            .post(plannerBody);

        if (plan.id === undefined) {
            console.log("retrieveOrCreatePlan(): failed to create plan id. response: ", plan);
            return;
        }
        else
            focusPlanId = plan.id;
    }

    // Get the existing bucket for "Focus Bucket", if not found, create it.
    res = await graphClient.api('/planner/plans/' + focusPlanId + '/buckets').get();

    var bucket = res.value.find(b => { return b.name === teamsContext.channelName + " Bucket"; });

    if ((res.value.length >= 1) && (bucket !== undefined)) {
        // found the bucket!
        focusBucketId = bucket.id;
        console.log("retrieveOrCreatePlan(): Bucket found! Bucket id: " + bucket.id + " Bucket name: " + bucket.name);
    }
    else {
        // create the bucket.
        var bucketBody = {
            "name": teamsContext.channelName + " Bucket",
            "planId": focusPlanId,
            "orderHint": " !"
        };

        bucket = await graphClient.api("planner/buckets")
            .post(bucketBody);

        if (bucket.id === undefined)
            console.log("SignIn(): Failed to create bucket. Response: ", bucket);
        else {
            console.log("SignIn(): Created bucket. Response: ", bucket);
            focusBucketId = bucket.id;
        }
    }
}

function getDateTimeOffset(dateStr) {
    var dt = new Date(dateStr),
        current_date = dt.getDate(),
        current_month = dt.getMonth() + 1,
        current_year = dt.getFullYear(),
        current_hrs = dt.getHours(),
        current_mins = dt.getMinutes(),
        current_secs = dt.getSeconds(),
        current_datetime;

    // Add 0 before date, month, hrs, mins or secs if they are less than 0
    current_date = current_date < 10 ? '0' + current_date : current_date;
    current_month = current_month < 10 ? '0' + current_month : current_month;
    current_hrs = current_hrs < 10 ? '0' + current_hrs : current_hrs;
    current_mins = current_mins < 10 ? '0' + current_mins : current_mins;
    current_secs = current_secs < 10 ? '0' + current_secs : current_secs;

    // Current datetime
    // String such as 2016-07-16T19:20:30
    current_datetime = current_year + '-' + current_month + '-' + current_date + 'T' + current_hrs + ':' + current_mins + ':' + current_secs;

    var timezone_offset_min = new Date().getTimezoneOffset(),
        offset_hrs = parseInt(Math.abs(timezone_offset_min / 60)),
        offset_min = Math.abs(timezone_offset_min % 60),
        timezone_standard;

    if (offset_hrs < 10)
        offset_hrs = '0' + offset_hrs;

    if (offset_min < 10)
        offset_min = '0' + offset_min;

    // Add an opposite sign to the offset
    // If offset is 0, it means timezone is UTC
    if (timezone_offset_min < 0)
        timezone_standard = '+' + offset_hrs + ':' + offset_min;
    else if (timezone_offset_min > 0)
        timezone_standard = '-' + offset_hrs + ':' + offset_min;
    else if (timezone_offset_min == 0)
        timezone_standard = 'Z';

    // Timezone difference in hours and minutes
    // String such as +5:30 or -6:00 or Z
    console.log(timezone_standard);
    console.log(current_datetime + timezone_standard);
    return current_datetime + timezone_standard;
}

async function getTasksAndAssignments() {

    if (focusBucketId)
        var taskData = await graphClient.api("planner/buckets/" + focusBucketId + "/tasks").get();
    else {
        console.log("getTasksAndAssignments(): no bucket id assigned!");
        return;
    }
    //var taskData = await client.api("planner/plans/" + gplanId + "/tasks").get();

    if (taskData === undefined) {
        console.log("getTasksAndAssignments(): task data is not available!");
        return;
    }

    gtaskData = taskData;

    //gtaskData.value.forEach(collatePeople);
    for (let i = 0; i < gtaskData.value.length; i++) {

        var peopleIds = Object.keys(gtaskData.value[i].assignments);
        for (let j = 0; j < peopleIds.length; j++) {

            if (gpeopleData[peopleIds[j]] !== undefined)
                break;
            else {
                let person = await graphClient.api("users/" + peopleIds[j]).get();
                if (person === undefined) {
                    console.log("getTasksAndAssignments(): couldn't find person: ", peopleIds[j]);
                    break;
                }
                gpeopleData[peopleIds[j]] = person;
            }
        }
    }
    if ((gGroupPeople !== null) && (gGroupPeople.value !== undefined))   {
        for (let i = 0; i < gGroupPeople.value.length; i++) {
            if (gpeopleData[gGroupPeople.value[i].id] !== undefined)
                break;
            else {
                let person = await graphClient.api("users/" + gGroupPeople.value[i].id).get();
                if (person === undefined) {
                    console.log("getTasksAndAssignments(): couldn't find person: ", gGroupPeople.value[i].id);
                    break;
                }
                gpeopleData[gGroupPeople.value[i].id] = person;

            }
        }
    }
}

function collatePeople(task) {
    var peopleIds = Object.keys(task.assignments);
    peopleIds.forEach(ppl => {
        if (gpeopleData[ppl] !== undefined)
            return;
        else {
            var person = graphClient.api("users/" + ppl).get().then(function (person) {
                if (person === undefined) {
                    console.log("initializePlanInfo(): couldn't find person: ", ppl);
                    return;
                }
                gpeopleData[ppl] = person;
            });
        }
    });
}

async function getAllTaskStatii() {

    var alltasks = await graphClient.api("planner/plans/" + focusPlanId + "/tasks").get();
    if (alltasks === undefined) {
        console.log("getAllTaskStatii(): failed to retrieve task information for plan: " + focusPlanId);
        return null;
    }
    var statArray = [];
    await Promise.all(alltasks.value.map(async (t) => {
        statArray.push({ id: t.id, percentComplete: t.percentComplete });
        console.log("getAllTaskStatii:", t);
    }));
    return statArray;
}

async function getTaskStatus(rectIndex, taskId) {
    var tasks = await graphClient.api("planner/tasks/" + taskId).get();
    if (tasks === undefined) {
        console.log("getTaskStatus(): failed to retrieve task information for id: " + taskId);
        return null;
    }
    var statArray = [];
    statArray.push({ rectId: rectIndex, id: tasks.id, percentComplete: tasks.percentComplete });
    return statArray;
}
