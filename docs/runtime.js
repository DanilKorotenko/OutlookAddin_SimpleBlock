let mailboxItem;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
    mailboxItem = Office.context.mailbox.item;
}

function validateMessage(event)
{
    console.log("Start validation stream");

    mailboxItem.notificationMessages.addAsync('NoSend',
        { type: 'errorMessage',
        message: 'Message blocked.' });
            
    event.completed({ allowEvent: false });
    return;
}
