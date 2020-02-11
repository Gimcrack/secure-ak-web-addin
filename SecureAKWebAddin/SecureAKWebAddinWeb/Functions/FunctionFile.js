Office.initialize = function () {
}

var encrypt_string = `<br>
<br>
<br>
[SecureAK Encryption String – Do Not Remove]`;

var decrypt_string = `<br>
<br>
<br>
[SecureAK Decryption String – Do Not Remove]`;

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

    // this will not work unless and until exchange on-prem supports the latest requirement set

    //Office.context.mailbox.item.internetHeaders.setAsync({ 'X-SAKAction': 'EncryptString' }, function (res) {
    //    console.log(res);
    //});

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, {}, function (body) {
        console.log(body.value);

        var msg = body.value;

        msg = encrypt_string + msg.replace(encrypt_string, '').replace(decrypt_string, '');


        Office.context.mailbox.item.body.setAsync(msg, {
            coercionType: Office.CoercionType.Html
        });
    });

    statusUpdate("icon16", "Message will be encrypted. Press Send to continue.");

    return event.completed();
}


function decryptAndSend(event) {

    // this will not work unless and until exchange on-prem supports the latest requirement set
   
    //Office.context.mailbox.item.internetHeaders.setAsync({ 'X-SAKAction': 'DecryptString' }, function (res) {
    //    console.log(res);
    //});

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, {}, function (body) {
        console.log(body.value);

        var msg = body.value;

        msg = decrypt_string + msg.replace(encrypt_string, '').replace(decrypt_string, '');


        Office.context.mailbox.item.body.setAsync(msg, {
            coercionType: Office.CoercionType.Html
        });
    });

    statusUpdate("icon16", "Message will be decrypted. Press Send to continue.");

    return event.completed();
}
