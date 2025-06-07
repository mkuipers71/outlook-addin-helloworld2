(function(){
    'use strict';

    Office.onReady(function(info) {
        if (info.host === Office.HostType.Outlook) {
            document.getElementById("export-to-provim").onclick = exportToProvim;
            loadEmailInfo();
        }
    });

    function loadEmailInfo() {
        // Get a reference to the current message
        var item = Office.context.mailbox.item;

        // Display basic info about the item
        document.getElementById("email-subject").innerHTML = "<strong>Subject:</strong> " + (item.subject || "No subject");
        
        if (item.from) {
            document.getElementById("email-from").innerHTML = "<strong>From:</strong> " + item.from.displayName + " (" + item.from.emailAddress + ")";
        } else {
            document.getElementById("email-from").innerHTML = "<strong>Item type:</strong> " + item.itemType;
        }

        // Add date if available
        if (item.dateTimeCreated) {
            var date = new Date(item.dateTimeCreated);
            document.getElementById("email-date").innerHTML = "<strong>Date:</strong> " + date.toLocaleString();
        } else {
            document.getElementById("email-date").innerHTML = "";
        }

        console.log("Email info loaded");
    }

    function exportToProvim() {
        var item = Office.context.mailbox.item;
        
        // Show the "not implemented" message
        Office.context.ui.displayDialogAsync(
            'data:text/html,<html><head><title>Provim Export</title><style>body{font-family:"Segoe UI",sans-serif;padding:30px;text-align:center;background:#f3f2f1;}h2{color:#0078d4;margin-bottom:20px;}p{margin:10px 0;font-size:14px;}.email-info{background:white;padding:15px;border-radius:8px;margin:20px 0;text-align:left;}.export-button{background:#0078d4;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;margin:10px;}button{background:#6c757d;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;}</style></head><body><h2>🔄 Export to Provim</h2><p>This functionality is not implemented yet.</p><div class="email-info"><strong>Email Details:</strong><br><strong>Subject:</strong> ' + (item.subject || 'No subject') + '<br><strong>Type:</strong> ' + item.itemType + '</div><p>This feature will allow you to:</p><ul style="text-align:left;"><li>Export email content to Provim</li><li>Link emails to Provim projects</li><li>Track email communications</li></ul><button onclick="window.close();">Close</button></body></html>',
            { height: 60, width: 60 },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    // Fallback to simple alert if dialog fails
                    alert('Export to Provim\n\nThis functionality is not implemented yet.\n\nEmail: ' + (item.subject || 'No subject'));
                }
            }
        );

        console.log("Export to Provim clicked - showing placeholder message");
    }

    // Make function available globally (for potential use from manifest button)
    window.exportToProvim = exportToProvim;
})();