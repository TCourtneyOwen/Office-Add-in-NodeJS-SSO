"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const path = require("path");
exports.addSecretCommandPath = path.resolve(`${__dirname}/scripts/addAppSecret.ps1`);
exports.azRestAddSecretCommandPath = path.resolve(`${__dirname}/scripts/azRestAddSecret.txt`);
exports.azRestAddTenantReplyUrlsCommandPath = path.resolve(`${__dirname}/scripts/azRestAddTenantReplyUrls.txt`);
exports.azRestGetOrganizationDetailsCommandPath = path.resolve(`${__dirname}/scripts/azRestGetOrganizationDetails.txt`);
exports.azRestGetTenantAdminMembershipCommandPath = path.resolve(`${__dirname}/scripts/azRestGetTenantAdminMembership.txt`);
exports.azRestGetTenantRolesCommandPath = path.resolve(`${__dirname}/scripts/azRestGetTenantRoles.txt`);
exports.azCliInstallCommandPath = path.resolve(`${__dirname}/scripts/azCliInstallCmd.ps1`);
exports.azRestAppCreateCommandPath = path.resolve(`${__dirname}/scripts/azRestAppCreateCmd.txt`);
exports.azRestSetIdentifierUriCommmandPath = path.resolve(`${__dirname}/scripts/azRestSetIdentifierUri.txt`);
exports.azRestSetSigninAudienceCommandPath = path.resolve(`${__dirname}/scripts/azSetSignInAudienceCmd.txt`);
exports.fallbackAuthDialogFilePath = path.resolve(`${process.cwd()}/src/public/javascripts/fallbackAuthDialog.js`);
exports.getInstalledAppsPath = path.resolve(`${__dirname}/scripts/getInstalledApps.ps1`);
exports.getSecretCommandPath = path.resolve(`${__dirname}/scripts/getAppSecret.ps1`);
exports.manifestPath = path.resolve(`${process.cwd()}/manifest.xml`);
exports.ssoDataFilePath = path.resolve(`${process.cwd()}/.ENV`);
//# sourceMappingURL=defaults.js.map