/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("displayDialogAsyncButton").onclick = openDialog;
  }
});

let dialog; // Declare dialog as global for use in later functions.
let accessToken;

function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://luspin.github.io/2407230050000622/dialogWindow.html',
    { height: 60, width: 30, promptBeforeOpen: false },
    function (asyncResult) {
      dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        processMessage(arg);
      });
    }
  );
}

function processMessage(arg) {
  // const messageFromDialog = JSON.parse(arg.message.slice(1, -1).replace(/\\"/g, '"'));

  dialog.close();

  /*
  if (messageFromDialog.messageType === "dialogClosed") {
    document.getElementById("dialogResultText").innerHTML = "Result: " + messageFromDialog.messageType;
    dialog.close();
  }
    */

  /*
  if (messageFromDialog.messageType === "userAuthenticated") {
    document.getElementById("dialogResultText").innerHTML = "Hello: " + messageFromDialog.displayName;
    accessToken = messageFromDialog.accessToken;
    dialog.close();
  }
    */
}
