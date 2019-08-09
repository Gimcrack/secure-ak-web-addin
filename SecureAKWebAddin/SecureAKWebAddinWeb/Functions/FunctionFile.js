Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function encryptAndSend(event) {


    Office.context.mailbox.item.internetHeaders.setAsync({ 'X-SAKAction': 'EncryptString' }, function (res) {
        console.log(res);
    });

    statusUpdate("icon16", "Message will be encrypted. Press Send to continue.");

    return event.completed();
}


function decryptAndSend(event) {
   
    Office.context.mailbox.item.internetHeaders.setAsync({ 'X-SAKAction': 'DecryptString' }, function (res) {
        console.log(res);
    });

    statusUpdate("icon16", "Message will be decrypted. Press Send to continue.");

    return event.completed();
}