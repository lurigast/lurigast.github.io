Office.initialize = function (reason) {
    
}

Office.onReady(() => {
    //Initiera Office.js
});

function BuildXMLRequestForRoomName(roomName) {

    var result =
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
<soap:Header>
<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />
</soap:Header>
<soap:Body>
    <ResolveNames xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
                  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                  ReturnFullContactData="true">
      <UnresolvedEntry>`+ roomName + `</UnresolvedEntry>
    </ResolveNames>
  </soap:Body>
</soap:Envelope>`
    return result;

};

function sendRequest(roomName) {
    var request = BuildXMLRequestForRoomName(roomName);

    return new Promise((resolve, reject) => {
        Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                console.log(result.value);
                const parsedValue = $.parseXML(result.value);
                let officeLocation = parsedValue.getElementsByTagName("t:OfficeLocation")[0].textContent;
                resolve(officeLocation);
            } else {
                reject(result.error);
            }
        });
    });
}

function SetLocationToAppointmentBody(LocationToBody) {

    let bodyFormat;
    let parsedText = parseHyperlinks(LocationToBody);
    console.log("parsedText: " + parsedText);

    Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
            return;
        }
        console.log("bodyFormat: " + asyncResult.value);
        if (asyncResult.value === 'html') {
            bodyFormat = "Office.CoercionType.Html";
        } else {
            bodyFormat = "Office.CoercionType.Text";
        }
        console.log("bodyFormat: " + bodyFormat);
    });

    Office.context.mailbox.item.body.prependAsync(parsedText, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
            return;
        }
    });
}

function parseHyperlinks(text) {
    var urlRegex = /\((https?:\/\/[^)]+)\)/g;
    return text.replace(urlRegex, '<a href="$1">$1</a>');
}

function addLocationToAppointmentBody(event) {

    var item = Office.context.mailbox.item;

    item.location.getAsync({ asyncContext: event }, (result) => {
        let event = result.asyncContext;
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Action failed with message ${result.error.message}`);
            event.completed({ allowEvent: false, errorMessage: "Failed to get the appointment's location." });
            return;
        }
        if (result.value === "") {
            item.notificationMessages.addSync("locationEmpty", {
                type: "errorMessage",
                message: "Please enter a location for the appointment."
            });
            event.completed({ allowEvent: false, errorMessage: "Don't forget to add a meeting location." });
            return;
        }

        console.log(`Appointment location: ${result.value}`);
        sendRequest(result.value).then((officeLocation) => {
            console.log("Office Location: ", officeLocation);
            if (!officeLocation.includes("https://")) {
                console.log("This room has no URL location. Contact IT support.")
                event.completed({ allowEvent: false, errorMessage: "Room has no URL." });
                return;
            }
            if (officeLocation === "") {
                console.log("This room has no data in its location field. Contact IT support");
                event.completed({ allowEvent: false, errorMessage: "Room has no containing data in its location attribute." });
                return;
            }

            SetLocationToAppointmentBody(officeLocation);
            event.completed({ allowEvent: true });
        }).catch((error) => {
            console.error("An error occured:", error);
        });
    })
};

Office.actions.associate("addLocationToAppointmentBody", addLocationToAppointmentBody)