const chalk = require('chalk');
const childProcess = require('child_process');
const defaults = require('./defaults');
require('dotenv').config();
const fs = require('fs');
const manifest = require('office-addin-manifest');
const ssoDataSettings = require('./ssoAppDataSetttings');

configureSSOApplication();

async function configureSSOApplication() {
    // Check to see if Azure CLI is installed.  If it isn't installed then install it
    const cliInstalled = await azureCliInstalled();
    if (!cliInstalled) {
        console.log("Azure CLI is not installed.  Installing now before proceeding");
        await installAzureCli();
        console.log('Please close your command shell, reopen and run configure-sso again.  This is neccessary to register the path to the Azure CLI');
        return;
    }

    const userJson = await logIntoAzure();
    if (Object.keys(userJson).length >= 1) {
        console.log('Login was successful!');
        const manifestInfo = await manifest.readManifestFile(defaults.manifestPath);

        // Register application
        const applicationJson = await createNewApplication(manifestInfo.displayName, userJson);

        // Write application data to manifest and .ENV file
        await ssoDataSettings.writeApplicationData(applicationJson.appId);

        // Log out of Azure
        await logoutAzure();

        // Output application definition to console
        console.log(applicationJson);
    }
    else {
        throw new Error(`Login to Azure did not succeed.`);
    }
}

async function createNewApplication(ssoAppName, userJson) {
    try {
        console.log('Registering new application in Azure');
        let azRestCommand = await fs.readFileSync(defaults.azRestAppCreateCommandPath, 'utf8');
        const reName = new RegExp('{SSO-AppName}', 'g');
        const rePort = new RegExp('{PORT}', 'g');
        azRestCommand = azRestCommand.replace(reName, ssoAppName).replace(rePort, process.env.PORT);

        const applicationJson = await promiseExecuteCommand(azRestCommand, true /* returnJson */);

        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
            // Set application IdentifierUri
            await setIdentifierUri(applicationJson);

            // Set application sign-in audience
            await setSignInAudience(applicationJson);

            // Grant admin consent for application
            if (await isUserTenantAdmin(userJson)){
                await grantAdminContent(applicationJson);
                await setTenantReplyUrls(applicationJson);
            } else {
                console.log(chalk.yellow("You are not a tenant admin so you cannot grant admin consent for your application.  Contact your tenant admin to grant consent"));
            }

            // Create an application secret and add to the credential store
            const secret = await setApplicationSecret(applicationJson);
            console.log(chalk.green(`App secret is ${secret}`));
            ssoDataSettings.addSecretToCredentialStore(ssoAppName, secret.secretText);

            return applicationJson;
        } else {
            console.log(chalk.red("Failed to register application"));
            return undefined;
        }

    } catch (err) {
        throw new Error(`Unable to register new application ${ssoAppName}. \n${err}`);
    }
}

async function applicationReady(applicationJson) {
    try {
        const azRestCommand = `az ad app show --id ${applicationJson.appId}`
        const appJson = await promiseExecuteCommand(azRestCommand, true /* returnJson */, true /* expectError */);
        return appJson !== "";
    } catch (err) {
        throw new Error(`Unable to get application info for ${applicationJson.displayName}. \n${err}`);
    }
}

async function grantAdminContent(applicationJson) {
    try {
        console.log('Granting admin consent');
        // Check to see if the application is available before granting admin consent
        let appReady = false;
        while (appReady === false) {
            appReady = await applicationReady(applicationJson);
        }
        let azRestCommand = `az ad app permission admin-consent --id ${applicationJson.appId}`;
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set grant admin consent for ${applicationJson.displayName}. \n${err}`);
    }
}

async function azureCliInstalled() {
    try {
        switch (process.platform) {
            case "win32":
                const appsInstalledWindowsCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getInstalledAppsPath}"`;
                const appsWindows = await promiseExecuteCommand(appsInstalledWindowsCommand);
                return appsWindows.filter(app => app.DisplayName === 'Microsoft Azure CLI').length > 0
            case "darwin":
                const appsInstalledMacCommand = 'brew list';
                const appsMac = await promiseExecuteCommand(appsInstalledMacCommand, false /* returnJson */);
                return appsMac.includes('azure-cli');;;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

async function installAzureCli() {
    try {
        console.log("Downloading and installing Azure CLI - this could take a minute or so");
        switch (process.platform) {
            case "win32":
                const windowsCliInstallCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.azCliInstallCommandPath}"`;
                await promiseExecuteCommand(windowsCliInstallCommand, false /* returnJson */);
                break;
            case "darwin": // macOS
                const macCliInstallCommand = 'brew update && brew install azure-cli';
                await promiseExecuteCommand(macCliInstallCommand, false /* returnJson */);
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

async function isUserTenantAdmin(userInfo) {
    console.log("Check if logged in user is a tenant admin");
    let azRestCommand = fs.readFileSync(defaults.azRestGetTenantRolesCommandPath, 'utf8');
    const tenantRoles = await promiseExecuteCommand(azRestCommand);
    let tenantAdminId = '';
    tenantRoles.value.forEach(item => {
        if (item.displayName === "Company Administrator") {
            tenantAdminId = item.id;
        }
    });

    azRestCommand = fs.readFileSync(defaults.azRestGetTenantAdminMembershipCommandPath, 'utf8');
    azRestCommand = azRestCommand.replace('<TENANT-ADMIN-ID>', tenantAdminId);
    const tenantAdmins = await promiseExecuteCommand(azRestCommand);
    let isTenantAdmin = false;
    tenantAdmins.value.forEach(item => {
        if (item.userPrincipalName === userInfo[0].user.name) {
            isTenantAdmin = true;
        }
    });
    return isTenantAdmin;
}

async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    let userJson = await promiseExecuteCommand('az login --allow-no-subscriptions');
    if (Object.keys(userJson).length < 1) {
        // Try alternate login
        logoutAzure();
        userJson = await promiseExecuteCommand('az login');
    }
    return userJson
}

async function logoutAzure() {
    console.log('Logging out of Azure now');
    return await promiseExecuteCommand('az logout');
}

async function promiseExecuteCommand(cmd, returnJson = true, expectError = false) {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, { maxBuffer: 1024 * 102400 }, async (err, stdout, stderr) => {
                if (err && !expectError) {
                    console.log(stderr);
                    reject(stderr);
                }
                
                let results = stdout;
                if (results !== '' && returnJson) {
                    results = JSON.parse(results);
                }
                resolve(results);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setApplicationSecret(applicationJson) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand = await fs.readFileSync(defaults.azRestAddSecretCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        const secretJson = await promiseExecuteCommand(azRestCommand);
        return secretJson;
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setIdentifierUri(applicationJson) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand = await fs.readFileSync(defaults.azRestSetIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId).replace('{PORT}', process.env.PORT);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setSignInAudience(applicationJson) {
    try {
        console.log('Setting signin audience');
        let azRestCommand = await fs.readFileSync(defaults.azRestSetSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set signInAudience for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setTenantReplyUrls(applicationJson) {
    try {

        let servicePrinicipaObjectlId = "";
        let setReplyUrls = true;
        const sharePointServiceId = "57fb890c-0dab-4253-a5e0-7188c88b2bb4";

        // Get tenant name and construct SharePoint SSO reply urls with tenant name
        let azRestCommand = fs.readFileSync(defaults.azRestGetOrganizationDetailsCommandPath, 'utf8');
        const tenantDetails = await promiseExecuteCommand(azRestCommand);
        const tenantName = tenantDetails.value[0].displayName
        const oneDriveReplyUrl = `https://${tenantName}-my.sharepoint.com/_forms/singlesignon.aspx`;
        const sharePointReplyUrl = `https://${tenantName}.sharepoint.com/_forms/singlesignon.aspx`;

        // Get service principals for tenant
        azRestCommand = 'az ad sp list --all';
        const servicePrincipals = await promiseExecuteCommand(azRestCommand);

        // Check if SharePoint redirects are set for SharePoint principal
        for (let item of servicePrincipals) {
            if (item.appId === sharePointServiceId) {
                servicePrinicipaObjectlId = item.objectId;
                if (item.replyUrls.length === 0) {
                    break;
                    // if there are reply urls set, then we need to see if the SharePoint SSO reply urls are already set
                } else {
                    for (let url of item.replyUrls) {
                        if (url === oneDriveReplyUrl || url === sharePointReplyUrl) {
                            setReplyUrls = false;
                            break;
                        }
                    }
                }

            }
        }

        if (setReplyUrls) {
            console.log('Setting SharePoint reply urls for tenant');
            azRestCommand = fs.readFileSync(defaults.azRestAddTenantReplyUrlsCommandPath, 'utf8');
            const reName = new RegExp('<TENANT-NAME>', 'g');
            azRestCommand = azRestCommand.replace(reName, tenantName).replace('<SP-OBJECTID>', servicePrinicipaObjectlId);
            await promiseExecuteCommand(azRestCommand);
        } else {
            console.log("SharePoint reply urls already set");
        }

    } catch (err) {
        throw new Error(`Unable to set tenant reply urls for ${applicationJson.displayName}. \n${err}`);
    }
}
