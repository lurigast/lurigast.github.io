﻿Office.initialize = function (reason) {
    
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

//    var result = `<?xml version="1.0" encoding="utf-8"?>
//<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
//<soap:Header>
//<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />
//</soap:Header>
//<soap:Body>
//<ResolveNames xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ReturnFullContactData="true">
//<UnresolvedEntry>`+ roomName + `</UnresolvedEntry>
//</ResolveNames>
//</soap:Body>
//</soap:Envelope>`;
    console.log(result);
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

    Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
            return;
        }
        bodyFormat = asyncResult.value;
        console.log("bodyFormat: " + bodyFormat);
    });

    Office.context.mailbox.item.body.prependAsync(LocationToBody, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
            return;
        }
    });
}

function addLocationToAppointmentBody(event) {

    var item = Office.context.mailbox.item;

    item.location.getAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Action failed with message ${result.error.message}`);
            return;
        }
        console.log(`Appointment location: ${result.value}`);
        sendRequest(result.value).then((officeLocation) => {
            console.log("Office Location: ", officeLocation),
                SetLocationToAppointmentBody(officeLocation);
            event.completed();
        }).catch((error) => {
            console.error("An error occured:", error);
        });
    })
};

Office.actions.associate("addLocationToAppointmentBody", addLocationToAppointmentBody)
