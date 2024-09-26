// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="./App.js" />

(function () {
    "use strict";
  
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Outlook) {
          // Check if this is an appointment form
          if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            // When an appointment is being scheduled, open the pop-up
            showInputPopup();
          }
        }
      });
    
      
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
  
            $('#insertText').click(insertText);
        });
    };
    
    function insertText() {
      var textToInsert1 = $('#textToInsert1').val();
      var textToInsert2 = $('#textToInsert2').val();

      var textToInsert = "<br/> Deal Id: " +textToInsert1+" Client CRDS Id: "+textToInsert2
      
      // Insert as plain text (CoercionType.Text)
      Office.context.mailbox.item.body.setSelectedDataAsync(
        textToInsert, 
        { coercionType: Office.CoercionType.Html }, 
        function (asyncResult) {
          // Display the result to the user
          if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
            //if(textToInsert == 'internal') {
              Office.addin.hide();
            
          }
          else {
            app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
          }
        });
    }

      function showInputPopup() {
        Office.context.ui.displayDialogAsync('https://localhost:3000/commands.html', { height: 20, width: 30 }, function (asyncResult) {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
            processDialogMessage(arg);
        });
        });
      }
      
      function processDialogMessage(arg) {
        const inputData = arg.message; // Capture input from the user
        console.log('User input received: ', inputData);
        //insertText(inputData)
        // Insert as plain text (CoercionType.Text)
      Office.context.mailbox.item.body.setSelectedDataAsync(
        "Appointment Classification: " + inputData, 
        { coercionType: Office.CoercionType.Text }, 
        function (asyncResult) {
          // Display the result to the user
          if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            
            if(inputData == 'Internal') {
              Office.context.ui.closeContainer();
            }
            app.showNotification("Success", "\"" + inputData + "\" inserted successfully.");
          }
          else {
            app.showNotification("Error", "Failed to insert \"" + inputData + "\": " + asyncResult.error.message);
          }
        });
        // You can now use the user input (e.g., store it, attach it to the appointment, etc.)
      }

  }
  
  )();

  function insertDealDetails() {
    Office.context.ui.displayDialogAsync("https://localhost:3000/taskpane.html", { height: 50, width: 50 },
    (asyncResult) => {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
           // dialog.close();
            insertText();
        });
    }
);
}

function onAppointmentSend(event) {
    Office.context.mailbox.item.notificationMessages.addAsync("prompt", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Please enter the user name:",
        icon: "icon-16",
        persistent: true
    });
  
    // Show a dialog to get the user name
    Office.context.ui.displayDialogAsync('https://localhost:3000/taskpane.html', { height: 30, width: 20 }, function (asyncResult) {
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

  