import { commands, ExtensionContext, window } from 'vscode';
import { view } from './constants';
import { AppRegDataProvider } from './data/applicationRegistration';
import { GraphClient } from './clients/graph';
import { AppReg } from './applicationRegistrations';
import { ApplicationService } from './services/application';
import { AppRolesService } from './services/appRoles';
import { KeyCredentialsService } from './services/keyCredentials';
import { OAuth2PermissionScopeService } from './services/oauth2PermissionScopes';
import { OwnerService } from './services/owner';
import { PasswordCredentialsService } from './services/passwordCredentials';
import { RedirectUriService } from './services/redirectUris';
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
	const appRolesService = new AppRolesService(dataProvider);
	const keyCredentialsService = new KeyCredentialsService(dataProvider);
	const oauth2PermissionScopeService = new OAuth2PermissionScopeService(dataProvider);
	const ownerService = new OwnerService(dataProvider);
	const passwordCredentialsService = new PasswordCredentialsService(dataProvider);
	const redirectUriService = new RedirectUriService(dataProvider);
	const requiredResourceAccessService = new RequiredResourceAccessService(dataProvider);
	const signInAudienceService = new SignInAudienceService(dataProvider);

	// Create a new instance of the Application Registrations class.
	const appReg = new AppReg(
		dataProvider,
		applicationService,
		appRolesService,
		keyCredentialsService,
		oauth2PermissionScopeService,
		ownerService,
		passwordCredentialsService,
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
	commands.registerCommand(`${view}.editAppIdUri`, app => appReg.editAppIdUri(app));
	commands.registerCommand(`${view}.removeAppIdUri`, app => appReg.removeAppIdUri(app));
	commands.registerCommand(`${view}.copyAppIdUri`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.copyRedirectUri`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.copyValue`, app => appReg.copyValue(app));
	commands.registerCommand(`${view}.addPasswordCredential`, credential => appReg.addPasswordCredential(credential));
	commands.registerCommand(`${view}.deletePasswordCredential`, credential => appReg.deletePasswordCredential(credential));
}