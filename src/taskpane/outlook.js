/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("submit-deal").style.display = "none";
    document.getElementById("dealForm").style.display = "none";
    
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

function onAppointmentSend(event) {
  Office.context.mailbox.item.notificationMessages.addAsync("prompt", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Please enter the user name:",
      icon: "icon-16",
      persistent: true
  });

  // Show a dialog to get the user name
  Office.context.ui.displayDialogAsync('https://localhost:3000/popup.html', { height: 30, width: 20 }, function (asyncResult) {
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
          // Handle the user input
          var userName = args.message;
          // Do something with the user name
          dialog.close();
          event.completed({ allowEvent: true });
      });
  });
}

export async function run() {
  /**
   * Insert your Outlook code here
   */

// Get a reference to the current message
document.getElementById("create-deal").style.display = "none";
document.getElementById("sideload-msg").style.display = "none";
document.getElementById("submit-deal").style.display = "block";
    document.getElementById("dealForm").style.display = "block";
const item = Office.context.mailbox.item;
console.log( "check mail suject"+item.subject);

Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
      // Do something with the result.
      console.log("Body: "+result.value);
      document.getElementById("dealDescription").value = result.value;
      //document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + result.value;
  });


//item.attachments //this is array
//item.dateTimeCreated
//item.dateTimeModified
//item.to //this is array
//item.sender:
  //displayName: "Girish Ahirrao"
  //emailAddress: "ahirrao.girish02@gmail.com"
  //recipientType: "externalUser"

//if any attachments
var outputString = "";

if (item.attachments.length > 0) {
    for (i = 0 ; i < item.attachments.length ; i++) {
        var attachment = item.attachments[i];
        outputString += "<BR>" + i + ". Name: ";
        outputString += attachment.name;
        outputString += "<BR>ID: " + attachment.id;
        outputString += "<BR>contentType: " + attachment.contentType;
        outputString += "<BR>size: " + attachment.size;
        outputString += "<BR>attachmentType: " + attachment.attachmentType;
        outputString += "<BR>isInline: " + attachment.isInline;
    }
}

console.log(outputString);

// Write message property value to the task pane
//document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
document.getElementById("dealName").value = item.subject;
document.getElementById("clientName").value = item.sender.displayName;
document.getElementById("dealDate").value = item.dateTimeCreated;
}
