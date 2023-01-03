import { commands, ExtensionContext, window, workspace, env, Disposable } from 'vscode';
import { view } from './constants';
import { AppRegDataProvider } from './data/applicationRegistration';
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
import { ActivityStatus } from './interfaces/activityStatus';

// Globals to deal with the list filter
let _filterCommand: string | undefined = undefined;
let _filterText: string | undefined = undefined;

// This method is called when the extension is activated.
export async function activate(context: ExtensionContext) {

	// Create a new instance of the GraphClient class.
	const graphClient = new GraphClient();

	// Create a new instance of the ApplicationDataProvider class.
	const dataProvider = new AppRegDataProvider(graphClient);

	// Create instances of the services.
	const applicationService = new ApplicationService(dataProvider, graphClient, context);
	const appRoleService = new AppRoleService(dataProvider, graphClient);
	const keyCredentialService = new KeyCredentialService(dataProvider, graphClient);
	const oauth2PermissionScopeService = new OAuth2PermissionScopeService(dataProvider, graphClient);
	const ownerService = new OwnerService(dataProvider, graphClient);
	const passwordCredentialService = new PasswordCredentialService(dataProvider, graphClient);
	const redirectUriService = new RedirectUriService(dataProvider, graphClient);
	const requiredResourceAccessService = new RequiredResourceAccessService(dataProvider, graphClient);
	const signInAudienceService = new SignInAudienceService(dataProvider, graphClient);

	workspace.onDidChangeConfiguration(event => {
		if (event.affectsConfiguration("applicationregistrations.showAllApplications") || event.affectsConfiguration("applicationregistrations.maximumApplicationsReturned")) {
			populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing application registrations..."));
		}
	});

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
	const errorHandler = (result: ActivityStatus) => {
		// Clear any status bar messages.
		result.statusBarHandle!.dispose();
	
		// Restore the original icon.
		result.treeViewItem!.iconPath = result.previousIcon;
		dataProvider.triggerOnDidChangeTreeData();
	
		// Display an error message.
		window.showErrorMessage(`An error occurred trying to complete your task: ${result.error!.message}.`);
	};

	// Define the populateTreeView function.
	const populateTreeView = async (statusBarHandle: Disposable | undefined = undefined) => {
        if (!dataProvider.isGraphClientInitialised) {
            dataProvider.initialiseGraphClient(statusBarHandle);
            return;
        }
        await dataProvider.renderTreeView("APPLICATIONS", statusBarHandle, _filterCommand);
    };

	// Define the filterTreeView function.
	const filterTreeView = async () => {
        // If the user is not authenticated then we don't want to do anything        
        if (!dataProvider.isGraphClientInitialised) {
            dataProvider.initialiseGraphClient();
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

	// Hook up the error handlers.
	applicationService.onError((result) => errorHandler(result));
	appRoleService.onError((result) => errorHandler(result));
	keyCredentialService.onError((result) => errorHandler(result));
	oauth2PermissionScopeService.onError((result) => errorHandler(result));
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

	commands.registerCommand(`${view}.signInToAzure`, () => graphClient.cliSignIn());
	commands.registerCommand(`${view}.addApp`, async () => await applicationService.add());
	commands.registerCommand(`${view}.refreshApps`, () => populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing application registrations...")));
	commands.registerCommand(`${view}.filterApps`, () => filterTreeView());

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Application Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.deleteApp`, async item => await applicationService.delete(item));
	commands.registerCommand(`${view}.renameApp`, async item => await applicationService.rename(item));
	commands.registerCommand(`${view}.viewAppManifest`, async item => await applicationService.viewManifest(item));
	commands.registerCommand(`${view}.copyClientId`, item => applicationService.copyClientId(item));
	commands.registerCommand(`${view}.openAppInPortal`, item => applicationService.openInPortal(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Role Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.addAppRole`, async item => await appRoleService.add(item));
	commands.registerCommand(`${view}.editAppRole`, async item => await appRoleService.edit(item));
	commands.registerCommand(`${view}.deleteAppRole`, async item => await appRoleService.delete(item));
	commands.registerCommand(`${view}.changeStateAppRole`, async item => await appRoleService.changeState(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Exposed API Scope Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.addExposedApiScope`, async item => await oauth2PermissionScopeService.add(item));
	commands.registerCommand(`${view}.editExposedApiScope`, async item => await oauth2PermissionScopeService.edit(item));
	commands.registerCommand(`${view}.deleteExposedApiScope`, async item => await oauth2PermissionScopeService.delete(item));
	commands.registerCommand(`${view}.changeStateExposedApiScope`, async item => await oauth2PermissionScopeService.changeState(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// App Id URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.editAppIdUri`, async item => await applicationService.editAppIdUri(item));
	commands.registerCommand(`${view}.removeAppIdUri`, async item => await applicationService.removeAppIdUri(item));
	commands.registerCommand(`${view}.copyAppIdUri`, item => copyValue(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Password Credentials Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.addPasswordCredential`, async item => await passwordCredentialService.add(item));
	commands.registerCommand(`${view}.deletePasswordCredential`, async item => await passwordCredentialService.delete(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Redirect URI Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.addRedirectUri`, async item => await redirectUriService.add(item));
	commands.registerCommand(`${view}.editRedirectUri`, async item => await redirectUriService.edit(item));
	commands.registerCommand(`${view}.deleteRedirectUri`, async item => await redirectUriService.delete(item));
	commands.registerCommand(`${view}.copyRedirectUri`, item => copyValue(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Sign In Audience Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.editAudience`, async item => await signInAudienceService.edit(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Owner Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.addOwner`, async item => await ownerService.add(item));
	commands.registerCommand(`${view}.removeOwner`, async item => await ownerService.remove(item));
	commands.registerCommand(`${view}.openUserInPortal`, item => ownerService.openInPortal(item));

	//////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Common Commands
	//////////////////////////////////////////////////////////////////////////////////////////////////////////

	commands.registerCommand(`${view}.copyValue`, item => copyValue(item));
}