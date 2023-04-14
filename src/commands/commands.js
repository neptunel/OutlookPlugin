


// 1. How to construct online meeting details.
// domain name can be changed based on which client you want to use (capi is for internal use only)
const domain = "https://capi.dolby.io/closed.beta.sp/";

const generateUniqueId = () => {
     // Generate a random number and convert it to a string
    const randomNumber = Math.floor(Math.random() * Date.now()).toString();
    // Generate a random 4-character string
    const randomString = Math.random().toString(36).substring(2, 6);
    // Concatenate the random number and random string
    return randomNumber + randomString;
};
const meetingID = generateUniqueId();

// const meetingLink = domain+ meetingID + "?name=" +`${userName}`;
const meetingLink = domain+ meetingID;
const QRGenerator = "https://chart.googleapis.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl=";
const newBody = '<br>' +
    `<a href= ${meetingLink} target="_blank">Join a Dolby.io Video Call</a>` +
    '<br><br>' +
    `Meeting ID: ${meetingID}` +
    '<br><br>' +
    'Want to test your video connection?' +
    '<br><br>' +
    '<a href="https://capi.dolby.io/closed.beta.sp/" target="_blank">Join a test meeting</a>' +
    '<br><br>'+
    'Scan the QR Code to join the meeting:'+
    '<br><br>'+
    `<img src= ${QRGenerator+meetingLink}>`+
    '<br><br>';
    

// 2. How to define and register a function command named `insertDolbyioMeeting` (referenced in the manifest)
//    to update the meeting body with the online meeting details.
function insertDolbyioMeeting(event) {
    // Get HTML body from the client. Uses updateBody to append at the end of existing body.
    // mailboxItem.body.getAsync("html",
    //     { asyncContext: event },
    //     function (getBodyResult) {
    //         if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
    //             updateBody(getBodyResult.asyncContext, getBodyResult.value);
    //         } else {
    //             console.error("Failed to get HTML body.");
    //             getBodyResult.asyncContext.completed({ allowEvent: false });
    //         }
    //     }
    // );
    // Sets body from scratch each time the add-in works.
        // Append new body to the existing body.
        mailboxItem.body.setAsync(newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    
        setLocation(event, meetingLink);
}

// 3. How to implement a supporting function `updateBody`
//    that appends the online meeting details to the current body of the meeting.
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
function setLocation(event,meetingLink){
    mailboxItem.location.setAsync(meetingLink,  
        { asyncContext: event },(result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        result.asyncContext.completed({ allowEvent: false });
        return;
        }
        console.log(`Successfully set location to ${meetingLink}`);
        result.asyncContext.completed({ allowEvent: true });
    });
}
let mailboxItem, userName;
// Office is ready.
Office.onReady(function () {
    mailboxItem = Office.context.mailbox.item;
    userName =  Office.context.mailbox.userProfile.displayName;
}
);
// Register the function.
Office.actions.associate("insertDolbyioMeeting", insertDolbyioMeeting);