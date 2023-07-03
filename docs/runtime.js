Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
}

async function validateMessage(event)
{
    console.log("Start message validation");

    event.completed({ allowEvent: true });
    return;
}

async function onMessageComposeHandler(event)
{
    console.log("onMessageComposeHandler");

    event.completed({ allowEvent: true });
    return;
}
