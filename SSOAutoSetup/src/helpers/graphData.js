import { writeFileNamesToOfficeDocument } from './document';
import { showMessage } from './message-helper';

export function makeGraphApiCall(accessToken) {
    $.ajax({
        type: "GET",
        url: "/getuserdata",
        headers: { "access_token": accessToken },
        cache: false
    }).done(function (response) {

        writeFileNamesToOfficeDocument(response)
            .then(function () {
                showMessage("Your data has been added to the document.");
            })
            .catch(function (error) {
                // The error from writeFileNamesToOfficeDocument will begin 
                // "Unable to add filenames to document."
                showMessage(error);
            });
    })
        .fail(function (errorResult) {
            // This error is relayed from `app.get('/getuserdata` in app.js file.
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
}