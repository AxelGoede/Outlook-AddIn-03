let item;

(function()  {
    function showAlert() {
        alert('Hallo Axel Goede - der erste Button wurde gedrückt!');
        
    }
    function showAlert02() {
        alert('Hallo Axel Goede - der zweite Button wurde gedrückt!');
    }
    function btnSetup() {
        document.getElementById('showButton').addEventListener('click',showAlert);
        document.getElementById('showButton02').addEventListener('click',showAlert02)
    }
    document.onreadystatechange=function() {
        if(document.readyState==="complete") {
            btnSetup();
        }
    }

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        setSubject();
    }
});

// Sets the subject of the item that the user is composing.
function setSubject() {
    // Customize the subject with today's date.
    const today = new Date();
    const subject = `Summary for ${today.toLocaleDateString()}`;
    
    item.subject.setAsync(
        subject,
        { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
                return;
            }

            /*
              The subject was successfully set.
              Run additional operations appropriate to your scenario and
              use the optionalVariable1 and optionalVariable2 values as needed.
            */
        });
}

}) ();

