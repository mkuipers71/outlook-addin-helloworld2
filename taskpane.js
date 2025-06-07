(function(){
    'use strict';

    Office.onReady(function(info) {
        if (info.host === Office.HostType.Outlook) {
            document.getElementById("get-item-class").onclick = getItemClass;
        }
    });

    function getItemClass() {
        // Get a reference to the current message
        var item = Office.context.mailbox.item;

        // Display basic info about the item
        document.getElementById("item-subject").innerHTML = "<strong>Subject:</strong> " + (item.subject || "No subject");
        
        if (item.from) {
            document.getElementById("item-from").innerHTML = "<strong>From:</strong> " + item.from.displayName + " (" + item.from.emailAddress + ")";
        } else {
            document.getElementById("item-from").innerHTML = "<strong>Item type:</strong> " + item.itemType;
        }

        console.log("Item class: " + item.itemClass);
        console.log("Item type: " + item.itemType);
    }
})();