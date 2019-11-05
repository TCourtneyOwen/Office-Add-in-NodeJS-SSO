/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. 
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */
Office.onReady(function(info) {
    $(document).ready(function() {
        $('#getGraphDataButton').click(getGraphData);
    });
});

let retryGetAccessToken = 0;

async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });
        let exchangeResponse = await getGraphToken(bootstrapToken);
        if (exchangeResponse.claims) {
            // Microsoft Graph requires an additional form of authentication. Have the Office host 
            // get a new token using the Claims string, which tells AAD to prompt the user for all 
            // required forms of authentication.
            let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
            exchangeResponse = await getGraphToken(mfaBootstrapToken);
        }
        
        if (exchangeResponse.error) {
            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            handleAADErrors(exchangeResponse);
        } 
        else 
        {
            // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
            // in the .fail callback of that call, not in the catch block below.
            makeGraphApiCall(exchangeResponse.access_token);
        }
    }
    catch(exception) {
        // The only exceptions caught here are exceptions in your code in the try block
        // and errors returned from the call of `getAccessToken` above.
        if (exception.code) { 
            handleClientSideErrors(exception);
        }
        else {
            showMessage("EXCEPTION: " + JSON.stringify(exception));
        } 
    }
}

async function getGraphToken(bootstrapToken) {
    let response = await $.ajax({type: "GET", 
		url: "/auth",
        headers: {"Authorization": "Bearer " + bootstrapToken }, 
        cache: false
    });
    return response;
}

function handleClientSideErrors(error) {
    switch (error.code) {

        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see this error
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
            break;
        case 13006:
            // Only seen in Office on the Web.
            showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
            break;
        case 13008:
            // Only seen in Office on the Web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
            break;
        case 13010:
            // Only seen in Office on the Web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            dialogFallback();
            break;
    }
}

function handleAADErrors(exchangeResponse) {
    // On rare occasions the bootstrap token is unexpired when Office validates it,
    // but expires by the time it is sent to AAD for exchange. AAD will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Retry the call of getAccessToken (no more than once). This time Office will return a 
    // new unexpired bootstrap token. 
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    else 
    {
        dialogFallback();
    }
}

function makeGraphApiCall(accessToken) {
    $.ajax({type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
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

function writeFileNamesToOfficeDocument(result) {

    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            switch (Office.context.host) {
                case "Excel":
                    writeFileNamesToWorksheet(result);
                    break;
                case "Word":
                    writeFileNamesToDocument(result);
                    break;
                case "PowerPoint":
                    writeFileNamesToPresentation(result);
                    break;
                default:
                    throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
            }
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error.toString()));
        }
    });    
}

function writeFileNamesToWorksheet(result) {

    return Excel.run(function (context) {
       const sheet = context.workbook.worksheets.getActiveWorksheet();

       let filenames = [];
       let i;
       for (i = 0; i < result.length; i++) {
           var innerArray = [];
           innerArray.push(result[i]);
           filenames.push(innerArray);
       }

       const rangeAddress = `B5:B${5 + (result.length-1)}`;
       const range = sheet.getRange(rangeAddress);
       range.values = filenames;
       range.format.autofitColumns();

       return context.sync();
   });
}

function writeFileNamesToDocument(result) {
     return Word.run(function (context) {
        const documentBody = context.document.body;
        for (let i = 0; i < result.length; i++) {
            documentBody.insertParagraph(result[i], "End");
        }

        return context.sync();
    });
}

function writeFileNamesToPresentation(result) {

    let fileNames = "";
    for (let i = 0; i < result.length; i++) {
        fileNames += result[i] + '\n';
    }

    Office.context.document.setSelectedDataAsync(
        fileNames,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        }
    );
}

function showMessage(text) {
    $('.welcome-body').hide();
    $('#message-area').show(); 
    $('#message-area').text(text);
 }

 var loginDialog;

function dialogFallback() {
	// We fall back to Dialog API for any error.
	// TODO: handle specific errors only?

    var url = "/fallbackAuthDialog.html"; 
	showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
function processMessage(arg) {

    console.log("Message received in processMessage: " + JSON.stringify(arg));
    let messageFromDialog = JSON.parse(arg.message);

        if (messageFromDialog.status === 'success') { 
            // We now have a valid access token.
            loginDialog.close();
            makeGraphApiCall(messageFromDialog.result);
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            showMessage(JSON.stringify(error.toString()));
        }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
	var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

	// height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
	Office.context.ui.displayDialogAsync(fullUrl,
		{ height: 60, width: 30 }, function (result) {
			console.log("Dialog has initialized. Wiring up events");
			loginDialog = result.value;
			loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
		});
}