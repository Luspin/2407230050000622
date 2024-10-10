Office.onReady(function () {
    document.getElementById("closeButton").onclick = closeButtonClick;
});

function closeButtonClick() {
    let messageObject_dialogClosed = { messageType: "dialogClosed" };
    let jsonMessage = JSON.stringify(messageObject_dialogClosed);
    Office.context.ui.messageParent(jsonMessage);
}