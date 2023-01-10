import { commands, window, workspace, env, Disposable, ExtensionContext } from "vscode";
import { VIEW_NAME } from "./constants";
import { AppRegTreeDataProvider } from "./data/app-reg-tree-data-provider";
import { AppRegItem } from "./models/app-reg-item";
import { GraphClient, escapeSingleQuotesForFilter } from "./clients/graph-client";
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
import { ActivityResult } from "./interfaces/activity-result";

// Values to hold the list filter
let filterCommand: string | undefined = undefined;
let filterText: string | undefined = undefined;

// Create a new instance of the GraphClient class.
const graphClient = new GraphClient();

// Create a new instance of the ApplicationTreeDataProvider class.
const treeDataProvider = new AppRegTreeDataProvider(graphClient);

// Create new instances of the services classes.
const applicationService = new ApplicationService(graphClient, treeDataProvider);
const appRoleService = new AppRoleService(graphClient, treeDataProvider);
const keyCredentialService = new KeyCredentialService(graphClient, treeDataProvider);
const oauth2PermissionScopeService = new OAuth2PermissionScopeService(graphClient, treeDataProvider);
const organizationService = new OrganizationService(graphClient);
const ownerService = new OwnerService(graphClient, treeDataProvider);
const passwordCredentialService = new PasswordCredentialService(graphClient, treeDataProvider);
const redirectUriService = new RedirectUriService(graphClient, treeDataProvider);
const requiredResourceAccessService = new RequiredResourceAccessService(graphClient, treeDataProvider);
const signInAudienceService = new SignInAudienceService(graphClient, treeDataProvider);

// This method is called when the extension is activated.
export const activate = async (context: ExtensionContext) => {

	workspace.onDidChangeConfiguration(event => {
		if (event.affectsConfiguration("applicationregistrations.showOwnedApplicationsOnly")
			|| event.affectsConfiguration("applicationregistrations.maximumQueryApps")
			|| event.affectsConfiguration("applicationregistrations.maximumApplicationsShown")
			|| event.affectsConfiguration("applicationregistrations.useEventualConsistency")) {
			populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing Application Registrations"));
		}
	});

	// Hook up the error handlers.
	treeDataProvider.onError((result) => errorHandler(result));
	applicationService.onError((result) => errorHandler(result));
	appRoleService.onError((result) => errorHandler(result));
	keyCredentialService.onError((result) => errorHandler(result));
	oauth2PermissionScopeService.onError((result) => errorHandler(result));
	organizationService.onError((result) => errorHandler(result));
	ownerService.onError((result) => errorHandler(result));
	passwordCredentialService.onError((result) => errorHandler(result));
	redirectUriService.onError((result) => errorHandler(result));
	requiredResourceAccessService.onError((result) => errorHandler(result));
	signInAudienceService.onError((result) => errorHandler(result));

	// Hook up the complete handlers.
	applicationService.onComplete((result) => populateTreeView(result.statusBarHandle));
	appRoleService.onComplete((result) => populateTreeView(result.statusBarHandle));
	keyCredentialService.onComplete((result) => populateTreeView(result.statusBarHandle));
	oauth2PermissionScopeService.onComplete((result) => populateTreeView(result.statusBarHandle));
	ownerService.onComplete((result) => populateTreeView(result.statusBarHandle));
	passwordCredentialService.onComplete((result) => populateTreeView(result.statusBarHandle));
	redirectUriService.onComplete((result) => populateTreeView(result.statusBarHandle));
	requiredResourceAccessService.onComplete((result) => populateTreeView(result.statusBarHandle));
	signInAudienceService.onComplete((result) => populateTreeView(result.statusBarHandle));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Menu Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.signInToAzure`, async () => await graphClient.cliSignIn()));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addApp`, async () => await applicationService.add()));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.refreshApps`, async () => await populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing Application Registrations"))));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.filterApps`, async () => await filterTreeView()));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.tenantInfo`, async () => await organizationService.showTenantInformation()));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Application Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteApp`, async item => await applicationService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.renameApp`, async item => await applicationService.rename(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.viewAppManifest`, async item => await applicationService.viewManifest(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyClientId`, item => applicationService.copyClientId(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.openAppInPortal`, item => applicationService.openInPortal(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Role Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addAppRole`, async item => await appRoleService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAppRole`, async item => await appRoleService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteAppRole`, async item => await appRoleService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.enableAppRole`, async item => await appRoleService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.disableAppRole`, async item => await appRoleService.changeState(item, false)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// API Permission Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addSingleScopeToExisting`, async item => await requiredResourceAccessService.addToExisting(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addSingleScope`, async item => await requiredResourceAccessService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeSingleScope`, async item => await requiredResourceAccessService.remove(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeApiScopes`, async item => await requiredResourceAccessService.removeApi(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exposed API Scope Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addExposedApiScope`, async item => await oauth2PermissionScopeService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editExposedApiScope`, async item => await oauth2PermissionScopeService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteExposedApiScope`, async item => await oauth2PermissionScopeService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.enableExposedApiScope`, async item => await oauth2PermissionScopeService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.disableExposedApiScope`, async item => await oauth2PermissionScopeService.changeState(item, false)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Id URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAppIdUri`, async item => await applicationService.editAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeAppIdUri`, async item => await applicationService.removeAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyAppIdUri`, item => copyValue(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Password Credentials Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addPasswordCredential`, async item => await passwordCredentialService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deletePasswordCredential`, async item => await passwordCredentialService.delete(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Key Credentials Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.uploadKeyCredential`, async item => await keyCredentialService.upload(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteKeyCredential`, async item => await keyCredentialService.delete(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Redirect URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addRedirectUri`, async item => await redirectUriService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editRedirectUri`, async item => await redirectUriService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.deleteRedirectUri`, async item => await redirectUriService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyRedirectUri`, item => copyValue(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Sign In Audience Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.editAudience`, async item => await signInAudienceService.edit(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Owner Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.addOwner`, async item => await ownerService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.removeOwner`, async item => await ownerService.remove(item)));
	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.openUserInPortal`, item => ownerService.openInPortal(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Common Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${VIEW_NAME}.copyValue`, item => copyValue(item)));
};

// This method is called when your extension is deactivated.
export const deactivate = async () => {
	signInAudienceService.dispose();
	requiredResourceAccessService.dispose();
	redirectUriService.dispose();
	passwordCredentialService.dispose();
	ownerService.dispose();
	oauth2PermissionScopeService.dispose();
	keyCredentialService.dispose();
	appRoleService.dispose();
	applicationService.dispose();
	graphClient.dispose();
	treeDataProvider.dispose();
};

// Define the populateTreeView function.
const populateTreeView = async (statusBarHandle: Disposable | undefined = undefined) => {
	if (!treeDataProvider.isGraphClientInitialised) {
		await treeDataProvider.initialiseGraphClient(statusBarHandle);
		return;
	}
	await treeDataProvider.renderTreeView("APPLICATIONS", statusBarHandle, filterCommand);
};

// Define the filterTreeView function.
const filterTreeView = async () => {
	// If the user is not authenticated then we don't want to do anything        
	if (!treeDataProvider.isGraphClientInitialised) {
		await treeDataProvider.initialiseGraphClient();
		return;
	}

	// If the tree is currently updating then we don't want to do anything.
	if (treeDataProvider.isUpdating) {
		return;
	}

	// If the tree is currently empty then we don't want to do anything.
	if (treeDataProvider.isTreeEmpty) {
		return;
	}

	// Determine if eventual consistency is enabled.
	const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;

	// If eventual consistency is disabled then we cannot apply the filter
	if (useEventualConsistency === false) {
		window.showInformationMessage("The application list cannot be filtered when not using eventual consistency. Please enable this in user settings first.", "OK");
		return;
	}

	// Prompt the user for the filter text.
	const newFilter = await window.showInputBox({
		placeHolder: "Name starts with",
		prompt: "Filter applications by display name",
		value: filterText,
		ignoreFocusOut: true
	});

	// Escape has been hit so we don't want to do anything.
	if ((newFilter === undefined) || (newFilter === '' && newFilter === (filterText ?? ""))) {
		return;
	} else if (newFilter === '' && filterText !== '') {
		filterText = undefined;
		filterCommand = undefined;
		await populateTreeView(window.setStatusBarMessage("$(loading~spin) Loading Application Registrations"));
	} else if (newFilter !== '' && newFilter !== filterText) {
		// If the filter text is not empty then set the filter command and filter text.
		filterText = newFilter!;
		filterCommand = `startswith(displayName, \'${escapeSingleQuotesForFilter(newFilter)}\')`;
		await populateTreeView(window.setStatusBarMessage("$(loading~spin) Filtering Application Registrations"));
	}
};

// Define the copy value function.
const copyValue = (item: AppRegItem) => {
	env.clipboard.writeText(
		item.contextValue === "COPY"
			|| item.contextValue === "WEB-REDIRECT-URI"
			|| item.contextValue === "SPA-REDIRECT-URI"
			|| item.contextValue === "NATIVE-REDIRECT-URI"
			|| item.contextValue === "APPID-URI"
			? item.value!
			: item.children![0].value!);
};

// Define the error handler function.
const errorHandler = async (result: ActivityResult) => {

	if (result.statusBarHandle !== undefined) {
		// Clear any status bar messages.
		result.statusBarHandle!.dispose();
	}

	if (result.treeViewItem !== undefined) {
		// Restore the original icon.
		result.treeViewItem!.iconPath = result.previousIcon;
		treeDataProvider.triggerOnDidChangeTreeData(result.treeViewItem);
	}

	if (result.error !== undefined) {
		// Log the error.
		console.error(result.error);

		// Display an error message.
		window.showErrorMessage(`An error occurred trying to complete your task: ${result.error!.message}.`, "OK");
	}
};
