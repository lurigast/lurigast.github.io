Office.initialize = function (reason) {
    // Initiera Office.js
}

Office.onReady(() => {
    // Ladda Office.js
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

function onError(message) {

    const id = "errorMessage"
    const details = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: message
    };
    Office.context.mailbox.item.notificationMessages.addAsync(id, details);

}

function ParseHyperlinks(text) {
    var urlRegex = /\((https?:\/\/[^)]+)\)/g;
    return text.replace(urlRegex, '<a href="$1">$1</a>');
}

async function GetRoomNameInput() {
    console.log("GetRoomNameInput");
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.location.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                if (result.value) {
                    resolve(result.value);
                } else {
                    console.error(`Action failed with message ${result.error}`);
                    onError("There's no room selected.");
                    reject(error);
                }
            } else {
                console.error(`Action failed with message ${result.error.message}`);
                let error = new Error(result.error.message);
                let message = new Office.NotificationMessages.ErrorMessage('Error'.error.message);
                onError("Failed to get room in location.");
                reject(result.error);
            }
        });
    });
}

async function SendSOAPRequestToResolveRoomNameToGetLocation(roomName) {
    var request = BuildXMLRequestForRoomName(roomName);
    return new Promise((resolve, reject) => {
        Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                console.log(result.value);
                const parsedValue = $.parseXML(result.value);
                let officeLocation = parsedValue.getElementsByTagName("t:OfficeLocation")[0].textContent;
                resolve(officeLocation);
            } else {
                onError("Failed fetching data for room location.");
                reject(result.error);
            }
        });
    });
}

async function GetBodyType() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getTypeAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}

async function SetLocationToAppointmentBody(locationName) {
    let bodyType = await GetBodyType();
    console.log(GetBodyType());
    console.log("here i am" + bodyType);
    return new Promise((resolve, reject) => {
        console.log("locationName: " + locationName);

        if (bodyType == Office.CoercionType.Html) {
            let parsedLocationName = ParseHyperlinks(locationName);
            console.log("parsedLocationName: " + parsedLocationName);
            Office.context.mailbox.item.body.prependAsync(parsedLocationName, { coercionType: Office.CoercionType.Html }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(result.error);
                }
            });
        }
        if (bodyType === Office.CoercionType.Text) {
            Office.context.mailbox.item.body.prependAsync(locationName, { coercionType: Office.CoercionType.Text }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(result.error);
                }
            });
        }
    });
}

function HasUrlInName(locationName) {
    if (locationName.includes("https://") || locationName.includes("http://")) {
        return true;
    }
    else {
        return false;
    }
}

// * Main funktion som anropas. * //
async function addLocationToAppointmentBody(event) {
    try {
        let roomName = await GetRoomNameInput();
        console.log('roomName' + roomName);

        let officeLocation = await SendSOAPRequestToResolveRoomNameToGetLocation(roomName)
        console.log('Appointment Location', officeLocation);

        if (HasUrlInName(officeLocation) === true) {
            await SetLocationToAppointmentBody(officeLocation);
            event.completed({ allowEvent: true });
        } else {
            console.log('hmm no url?')
            onError("This room has no valid URL. Contact support if needed.");
            event.completed({ allowEvent: false, errorMessage: "Room has no URL." });
        }
    } catch (error) {
        console.error('Error: ' + JSON.stringify(error));
        event.completed({ allowEvent: false, errorMessage: error });
    }
};

Office.actions.associate("addLocationToAppointmentBody", addLocationToAppointmentBody)