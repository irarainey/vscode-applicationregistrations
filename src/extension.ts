import { commands, ExtensionContext, window } from 'vscode';
import { view } from './constants';
import { AppRegDataProvider } from './data/applicationRegistration';
import { GraphClient } from './clients/graph';
import { AppReg } from './applicationRegistrations';
import { ApplicationService } from './services/application';
import { AppRoleService } from './services/appRole';
import { KeyCredentialService } from './services/keyCredential';
import { OAuth2PermissionScopeService } from './services/oauth2PermissionScope';
import { OwnerService } from './services/owner';
import { PasswordCredentialService } from './services/passwordCredential';
import { RedirectUriService } from './services/redirectUri';
import { RequiredResourceAccessService } from './services/requiredResourceAccess';
import { SignInAudienceService } from './services/signInAudience';

// This method is called when the extension is activated.
export async function activate(context: ExtensionContext) {

	// Create a new instance of the GraphClient class.
	const graphClient = new GraphClient();

	// Create a new instance of the ApplicationDataProvider class.
	const dataProvider = new AppRegDataProvider(graphClient);

	// Create instances of the services.
	const applicationService = new ApplicationService(dataProvider, context);
	const appRoleService = new AppRoleService(dataProvider);
	const keyCredentialService = new KeyCredentialService(dataProvider);
	const oauth2PermissionScopeService = new OAuth2PermissionScopeService(dataProvider);
	const ownerService = new OwnerService(dataProvider);
	const passwordCredentialService = new PasswordCredentialService(dataProvider);
	const redirectUriService = new RedirectUriService(dataProvider);
	const requiredResourceAccessService = new RequiredResourceAccessService(dataProvider);
	const signInAudienceService = new SignInAudienceService(dataProvider);

	// Create a new instance of the Application Registrations class.
	const appReg = new AppReg(
		dataProvider,
		applicationService,
		appRoleService,
		keyCredentialService,
		oauth2PermissionScopeService,
		ownerService,
		passwordCredentialService,
		redirectUriService,
		requiredResourceAccessService,
		signInAudienceService
	);

	// Register the commands.
	commands.registerCommand(`${view}.signInToAzure`, () => graphClient.cliSignIn());
	commands.registerCommand(`${view}.refreshApps`, () => appReg.populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing Application Registrations...")));
	commands.registerCommand(`${view}.filterApps`, () => appReg.filterTreeView());
	commands.registerCommand(`${view}.addApp`, () => appReg.addApp());
	commands.registerCommand(`${view}.deleteApp`, app => appReg.deleteApp(app));
	commands.registerCommand(`${view}.renameApp`, app => appReg.renameApp(app));
	commands.registerCommand(`${view}.copyClientId`, app => appReg.copyClientId(app));
	commands.registerCommand(`${view}.openAppInPortal`, app => appReg.openAppInPortal(app));
	commands.registerCommand(`${view}.viewAppManifest`, app => appReg.viewAppManifest(app));
	commands.registerCommand(`${view}.editAudience`, app => appReg.editAudience(app));
	commands.registerCommand(`${view}.addOwner`, user => appReg.addOwner(user));
	commands.registerCommand(`${view}.removeOwner`, user => appReg.removeOwner(user));
	commands.registerCommand(`${view}.openUserInPortal`, user => appReg.openUserInPortal(user));
	commands.registerCommand(`${view}.addRedirectUri`, app => appReg.addRedirectUri(app));
	commands.registerCommand(`${view}.editRedirectUri`, app => appReg.editRedirectUri(app));
	commands.registerCommand(`${view}.deleteRedirectUri`, app => appReg.deleteRedirectUri(app));
	commands.registerCommand(`${view}.addAppRole`, app => appReg.addAppRole(app));
	commands.registerCommand(`${view}.editAppRole`, app => appReg.editAppRole(app));
	commands.registerCommand(`${view}.deleteAppRole`, app => appReg.deleteAppRole(app));
	commands.registerCommand(`${view}.changeStateAppRole`, app => appReg.changeStateAppRole(app));
	commands.registerCommand(`${view}.editAppIdUri`, app => appReg.editAppIdUri(app));
	commands.registerCommand(`${view}.removeAppIdUri`, app => appReg.removeAppIdUri(app));
	commands.registerCommand(`${view}.copyAppIdUri`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.copyRedirectUri`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.copyValue`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.addPasswordCredential`, credential => appReg.addPasswordCredential(credential));
	commands.registerCommand(`${view}.deletePasswordCredential`, credential => appReg.deletePasswordCredential(credential));
}