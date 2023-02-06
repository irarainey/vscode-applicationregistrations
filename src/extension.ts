import { commands, window, workspace, ExtensionContext } from "vscode";
import { VIEW_NAME } from "./constants";
import { GraphApiRepository } from "./repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "./data/app-reg-tree-data-provider";
import { ApplicationService } from "./services/application";
import { AppRoleService } from "./services/app-role";
import { KeyCredentialService } from "./services/key-credential";
import { OAuth2PermissionScopeService } from "./services/oauth2-permission-scope";
import { OrganizationService } from "./services/organization";
import { OwnerService } from "./services/owner";
import { PasswordCredentialService } from "./services/password-credential";
import { RedirectUriService } from "./services/redirect-uri";
import { RequiredResourceAccessService } from "./services/required-resource-access";
import { SignInAudienceService } from "./services/sign-in-audience";
import { copyValue } from "./utils/copy-value";
import { setStatusBarMessage } from "./utils/status-bar";
import { authenticate } from "./utils/cli-authentication";

// Create a new instance of the Graph Api Repository.
const graphRepository = new GraphApiRepository();

// Create a new instance of the ApplicationTreeDataProvider class.
const treeDataProvider = new AppRegTreeDataProvider(graphRepository);

// Create new instances of the services classes.
const applicationService = new ApplicationService(graphRepository, treeDataProvider);
const appRoleService = new AppRoleService(graphRepository, treeDataProvider);
const keyCredentialService = new KeyCredentialService(graphRepository, treeDataProvider);
const oauth2PermissionScopeService = new OAuth2PermissionScopeService(graphRepository, treeDataProvider);
const organizationService = new OrganizationService(graphRepository, treeDataProvider);
const ownerService = new OwnerService(graphRepository, treeDataProvider);
const passwordCredentialService = new PasswordCredentialService(graphRepository, treeDataProvider);
const redirectUriService = new RedirectUriService(graphRepository, treeDataProvider);
const requiredResourceAccessService = new RequiredResourceAccessService(graphRepository, treeDataProvider);
const signInAudienceService = new SignInAudienceService(graphRepository, treeDataProvider);

// Extension Activation
export async function activate(context: ExtensionContext) {

	// Hook up the configuration setting change handlers.
	workspace.onDidChangeConfiguration(async (event) => {
		if (event.affectsConfiguration("applicationregistrations.showOwnedApplicationsOnly")
			|| event.affectsConfiguration("applicationregistrations.maximumQueryApps")
			|| event.affectsConfiguration("applicationregistrations.maximumApplicationsShown")
			|| event.affectsConfiguration("applicationregistrations.useEventualConsistency")) {
			await treeDataProvider.render(setStatusBarMessage("Refreshing Application Registrations..."));
		}
	});

	// Auth Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.signIn`, async () => {}));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.signOut`, async () => {}));

	// Menu Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addApp`, async () => await applicationService.add()));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.refreshApps`, async () => await treeDataProvider.render(setStatusBarMessage("Refreshing Application Registrations..."))));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.filterApps`, async () => await treeDataProvider.filter()));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.tenantInfo`, async () => await organizationService.showTenantInformation()));

	// Application Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteApp`, async item => await applicationService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.renameApp`, async item => await applicationService.rename(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.viewAppManifest`, async item => await applicationService.viewManifest(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyClientId`, item => applicationService.copyClientId(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.showEndpoints`, item => applicationService.showEndpoints(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.openAppInPortal`, item => applicationService.openInPortal(item)));

	// App Role Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addAppRole`, async item => await appRoleService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAppRole`, async item => await appRoleService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAppRoleValue`, async item => await appRoleService.editValue(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.renameAppRole`, async item => await appRoleService.editValue(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteAppRole`, async item => await appRoleService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.enableAppRole`, async item => await appRoleService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.disableAppRole`, async item => await appRoleService.changeState(item, false)));

	// API Permission Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addSingleScopeToExisting`, async item => await requiredResourceAccessService.addToExisting(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addSingleScope`, async item => await requiredResourceAccessService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeSingleScope`, async item => await requiredResourceAccessService.remove(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeApiScopes`, async item => await requiredResourceAccessService.removeApi(item)));

	// Exposed API Scope Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.renameExposedApiScope`, async item => await oauth2PermissionScopeService.editValue(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editExposedApiScopeValue`, async item => await oauth2PermissionScopeService.editValue(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addExposedApiScope`, async item => await oauth2PermissionScopeService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editExposedApiScope`, async item => await oauth2PermissionScopeService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteExposedApiScope`, async item => await oauth2PermissionScopeService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.enableExposedApiScope`, async item => await oauth2PermissionScopeService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.disableExposedApiScope`, async item => await oauth2PermissionScopeService.changeState(item, false)));

	// App Id URI Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAppIdUri`, async item => await applicationService.editAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeAppIdUri`, async item => await applicationService.removeAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyAppIdUri`, item => copyValue(item)));

	// Password Credentials Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addPasswordCredential`, async item => await passwordCredentialService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deletePasswordCredential`, async item => await passwordCredentialService.delete(item)));

	// Key Credentials Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.uploadKeyCredential`, async item => await keyCredentialService.upload(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteKeyCredential`, async item => await keyCredentialService.delete(item)));

	// Redirect URI Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addRedirectUri`, async item => await redirectUriService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editRedirectUri`, async item => await redirectUriService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteRedirectUri`, async item => await redirectUriService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteAllRedirectUris`, async item => await redirectUriService.deleteAll(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyRedirectUri`, item => copyValue(item)));

	// Sign In Audience Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAudience`, async item => await signInAudienceService.edit(item)));

	// Owner Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addOwner`, async item => await ownerService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeOwner`, async item => await ownerService.remove(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.openUserInPortal`, item => ownerService.openInPortal(item)));

	// Common Commands
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyValue`, item => copyValue(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.signInToAzure`, async () => {
		await treeDataProvider.render(undefined, "AUTHENTICATING");
		const status = await authenticate();
		if (status === true) {
			await treeDataProvider.render(setStatusBarMessage("Loading Application Registrations..."), "AUTHENTICATED");
		} else {
			await treeDataProvider.render(undefined, "SIGN-IN");
		}
	}));

	// Register the tree data provider.
	window.registerTreeDataProvider(VIEW_NAME, treeDataProvider);

	// Set the tree view to initialising.
	await treeDataProvider.render(undefined, "INITIALISING");

	// Initialise the graph repository.
	const status = await graphRepository.initialise();

	// Were we able to get an access token?
	if (status !== true) {
		// If not then the user needs to authenticate.
		await treeDataProvider.render(undefined, "SIGN-IN");
	} else {
		// If so then the user is authenticated so load the tree view.
		await treeDataProvider.render(setStatusBarMessage("Loading Application Registrations..."));
	}
}

// Extension Deactivation
export function deactivate(): void {
	// Dispose the services, tree data provider, and graph client.
	signInAudienceService.dispose();
	requiredResourceAccessService.dispose();
	redirectUriService.dispose();
	passwordCredentialService.dispose();
	organizationService.dispose();
	ownerService.dispose();
	oauth2PermissionScopeService.dispose();
	keyCredentialService.dispose();
	appRoleService.dispose();
	applicationService.dispose();
	graphRepository.dispose();
	treeDataProvider.dispose();
}