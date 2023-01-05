import { commands, window, workspace, env, Disposable, ExtensionContext } from 'vscode';
import { view } from './constants';
import { AppRegTreeDataProvider } from './data/appRegTreeDataProvider';
import { AppRegItem } from './models/appRegItem';
import { GraphClient } from './clients/graph';
import { ApplicationService } from './services/application';
import { AppRoleService } from './services/appRole';
import { KeyCredentialService } from './services/keyCredential';
import { OAuth2PermissionScopeService } from './services/oauth2PermissionScope';
import { OwnerService } from './services/owner';
import { PasswordCredentialService } from './services/passwordCredential';
import { RedirectUriService } from './services/redirectUri';
import { RequiredResourceAccessService } from './services/requiredResourceAccess';
import { SignInAudienceService } from './services/signInAudience';
import { ActivityResult } from './interfaces/activityResult';

// Globals to deal with the list filter
let _filterCommand: string | undefined = undefined;
let _filterText: string | undefined = undefined;

// Create a new instance of the GraphClient class.
const _graphClient = new GraphClient();

// Create a new instance of the ApplicationTreeDataProvider class.
const _treeDataProvider = new AppRegTreeDataProvider(_graphClient);

// Create instances of the services.
const _applicationService = new ApplicationService(_treeDataProvider, _graphClient);
const _appRoleService = new AppRoleService(_treeDataProvider, _graphClient);
const _keyCredentialService = new KeyCredentialService(_treeDataProvider, _graphClient);
const _oauth2PermissionScopeService = new OAuth2PermissionScopeService(_treeDataProvider, _graphClient);
const _ownerService = new OwnerService(_treeDataProvider, _graphClient);
const _passwordCredentialService = new PasswordCredentialService(_treeDataProvider, _graphClient);
const _redirectUriService = new RedirectUriService(_treeDataProvider, _graphClient);
const _requiredResourceAccessService = new RequiredResourceAccessService(_treeDataProvider, _graphClient);
const _signInAudienceService = new SignInAudienceService(_treeDataProvider, _graphClient);

// This method is called when the extension is activated.
export async function activate(context: ExtensionContext) {

	workspace.onDidChangeConfiguration(event => {
		if (event.affectsConfiguration("applicationregistrations.showOwnedApplicationsOnly")
		|| event.affectsConfiguration("applicationregistrations.maximumQueryApps")
		|| event.affectsConfiguration("applicationregistrations.maximumApplicationsShown")
		|| event.affectsConfiguration("applicationregistrations.useEventualConsistency")) {
			populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing application registrations..."));
		}
	});

	// Hook up the error handlers.
	_treeDataProvider.onError((result) => errorHandler(result));
	_applicationService.onError((result) => errorHandler(result));
	_appRoleService.onError((result) => errorHandler(result));
	_keyCredentialService.onError((result) => errorHandler(result));
	_oauth2PermissionScopeService.onError((result) => errorHandler(result));
	_ownerService.onError((result) => errorHandler(result));
	_passwordCredentialService.onError((result) => errorHandler(result));
	_redirectUriService.onError((result) => errorHandler(result));
	_requiredResourceAccessService.onError((result) => errorHandler(result));
	_signInAudienceService.onError((result) => errorHandler(result));

	// Hook up the complete handlers.
	_applicationService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_appRoleService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_keyCredentialService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_oauth2PermissionScopeService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_ownerService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_passwordCredentialService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_redirectUriService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_requiredResourceAccessService.onComplete((result) => populateTreeView(result.statusBarHandle));
	_signInAudienceService.onComplete((result) => populateTreeView(result.statusBarHandle));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Menu Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.signInToAzure`, () => _graphClient.cliSignIn()));
	context.subscriptions.push(commands.registerCommand(`${view}.addApp`, async () => await _applicationService.add()));
	context.subscriptions.push(commands.registerCommand(`${view}.refreshApps`, () => populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing application registrations..."))));
	context.subscriptions.push(commands.registerCommand(`${view}.filterApps`, () => filterTreeView()));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Application Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.deleteApp`, async item => await _applicationService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.renameApp`, async item => await _applicationService.rename(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.viewAppManifest`, async item => await _applicationService.viewManifest(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.copyClientId`, item => _applicationService.copyClientId(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.openAppInPortal`, item => _applicationService.openInPortal(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Role Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.addAppRole`, async item => await _appRoleService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.editAppRole`, async item => await _appRoleService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.deleteAppRole`, async item => await _appRoleService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.enableAppRole`, async item => await _appRoleService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${view}.disableAppRole`, async item => await _appRoleService.changeState(item, false)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exposed API Scope Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.addExposedApiScope`, async item => await _oauth2PermissionScopeService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.editExposedApiScope`, async item => await _oauth2PermissionScopeService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.deleteExposedApiScope`, async item => await _oauth2PermissionScopeService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.enableExposedApiScope`, async item => await _oauth2PermissionScopeService.changeState(item, true)));
	context.subscriptions.push(commands.registerCommand(`${view}.disableExposedApiScope`, async item => await _oauth2PermissionScopeService.changeState(item, false)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Id URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.editAppIdUri`, async item => await _applicationService.editAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.removeAppIdUri`, async item => await _applicationService.removeAppIdUri(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.copyAppIdUri`, item => copyValue(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Password Credentials Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.addPasswordCredential`, async item => await _passwordCredentialService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.deletePasswordCredential`, async item => await _passwordCredentialService.delete(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Redirect URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.addRedirectUri`, async item => await _redirectUriService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.editRedirectUri`, async item => await _redirectUriService.edit(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.deleteRedirectUri`, async item => await _redirectUriService.delete(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.copyRedirectUri`, item => copyValue(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Sign In Audience Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.editAudience`, async item => await _signInAudienceService.edit(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Owner Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.addOwner`, async item => await _ownerService.add(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.removeOwner`, async item => await _ownerService.remove(item)));
	context.subscriptions.push(commands.registerCommand(`${view}.openUserInPortal`, item => _ownerService.openInPortal(item)));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Common Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	context.subscriptions.push(commands.registerCommand(`${view}.copyValue`, item => copyValue(item)));
}

// This method is called when your extension is deactivated.
export async function deactivate() {
	_signInAudienceService.dispose();
	_requiredResourceAccessService.dispose();
	_redirectUriService.dispose();
	_passwordCredentialService.dispose();
	_ownerService.dispose();
	_oauth2PermissionScopeService.dispose();
	_keyCredentialService.dispose();
	_appRoleService.dispose();
	_applicationService.dispose();
	_graphClient.dispose();
	_treeDataProvider.dispose();
}

// Define the populateTreeView function.
async function populateTreeView(statusBarHandle: Disposable | undefined = undefined) {
	if (!_treeDataProvider.isGraphClientInitialised) {
		_treeDataProvider.initialiseGraphClient(statusBarHandle);
		return;
	}
	await _treeDataProvider.renderTreeView("APPLICATIONS", statusBarHandle, _filterCommand);
};

// Define the filterTreeView function.
async function filterTreeView() {
	// If the user is not authenticated then we don't want to do anything        
	if (!_treeDataProvider.isGraphClientInitialised) {
		_treeDataProvider.initialiseGraphClient();
		return;
	}

	// Determine if eventual consistency is enabled.
	const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;

	// If eventual consistency is disabled then we cannot apply the filter
	if(useEventualConsistency === false) {
		window.showInformationMessage("The application list cannot be filtered when not using eventual consistency. Please enable this in user settings first.", "OK");
		return;
	}

	// Prompt the user for the filter text.
	const newFilter = await window.showInputBox({
		placeHolder: "Name starts with...",
		prompt: "Filter applications by display name",
		value: _filterText,
		ignoreFocusOut: true
	});

	// Escape has been hit so we don't want to do anything.
	if ((newFilter === undefined) || (newFilter === '' && newFilter === (_filterText ?? ""))) {
		return;
	} else if (newFilter === '' && _filterText !== '') {
		_filterText = undefined;
		_filterCommand = undefined;
		await populateTreeView(window.setStatusBarMessage("$(loading~spin) Loading application registrations..."));
	} else if (newFilter !== '' && newFilter !== _filterText) {
		// If the filter text is not empty then set the filter command and filter text.
		_filterText = newFilter!;
		_filterCommand = `startsWith(displayName, \'${newFilter}\')`;
		await populateTreeView(window.setStatusBarMessage("$(loading~spin) Filtering application registrations..."));
	}
};

// Define the copy value function.
function copyValue(item: AppRegItem) {
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
async function errorHandler(result: ActivityResult) {

	if (result.statusBarHandle !== undefined) {
		// Clear any status bar messages.
		result.statusBarHandle!.dispose();
	}

	if (result.treeViewItem !== undefined) {
		// Restore the original icon.
		result.treeViewItem!.iconPath = result.previousIcon;
		_treeDataProvider.triggerOnDidChangeTreeData(result.treeViewItem);
	}

	// Log the error.
	console.error(result.error);

	// Display an error message.
	window.showErrorMessage(`An error occurred trying to complete your task: ${result.error!.message}.`, "OK");
};
