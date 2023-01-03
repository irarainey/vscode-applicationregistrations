import { window, env, Disposable, workspace, ThemeIcon } from 'vscode';
import { AppRegDataProvider } from './data/applicationRegistration';
import { AppRegItem } from './models/appRegItem';
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
import { stat } from 'fs';

// This class is responsible for managing the application registrations tree view.
export class AppReg {

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // A private string to store the filter command.
    private _filterCommand?: string = undefined;

    // A private string to store the filter text.
    private _filterText?: string = undefined;

    // Private instances of our services
    private _applicationService: ApplicationService;
    private _appRoleService: AppRoleService;
    private _keyCredentialService: KeyCredentialService;
    private _oauth2PermissionScopeService: OAuth2PermissionScopeService;
    private _ownerService: OwnerService;
    private _passwordCredentialService: PasswordCredentialService;
    private _redirectUriService: RedirectUriService;
    private _requiredResourceAccessService: RequiredResourceAccessService;
    private _signInAudienceService: SignInAudienceService;

    // The constructor for the ApplicationRegistrations class.
    constructor(
        dataProvider: AppRegDataProvider,
        applicationService: ApplicationService,
        appRolesService: AppRoleService,
        keyCredentialsService: KeyCredentialService,
        oauth2PermissionScopeService: OAuth2PermissionScopeService,
        ownerService: OwnerService,
        passwordCredentialsService: PasswordCredentialService,
        redirectUriService: RedirectUriService,
        requiredResourceAccessService: RequiredResourceAccessService,
        signInAudienceService: SignInAudienceService) {
        this._dataProvider = dataProvider;
        this._applicationService = applicationService;
        this._appRoleService = appRolesService;
        this._keyCredentialService = keyCredentialsService;
        this._oauth2PermissionScopeService = oauth2PermissionScopeService;
        this._ownerService = ownerService;
        this._passwordCredentialService = passwordCredentialsService;
        this._redirectUriService = redirectUriService;
        this._requiredResourceAccessService = requiredResourceAccessService;
        this._signInAudienceService = signInAudienceService;

        this._appRoleService.onError((result) => this.errorHandler(result));
        this._oauth2PermissionScopeService.onError((result) => this.errorHandler(result));
        
        this._appRoleService.onComplete((result) => this.populateTreeView(result.statusBarHandle));
        this._oauth2PermissionScopeService.onComplete((result) => this.populateTreeView(result.statusBarHandle));

        workspace.onDidChangeConfiguration(event => {
            if (event.affectsConfiguration("applicationregistrations.showAllApplications") || event.affectsConfiguration("applicationregistrations.maximumApplicationsReturned")) {
                this.populateTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing application registrations..."));
            }
        });
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Tree View
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Populates the tree view with the applications.
    public async populateTreeView(statusBarHandle: Disposable | undefined = undefined): Promise<void> {
        if (!this._dataProvider.isGraphClientInitialised) {
            this._dataProvider.initialiseGraphClient(statusBarHandle);
            return;
        }
        await this._dataProvider.renderTreeView("APPLICATIONS", statusBarHandle, this._filterCommand);
    }

    // Filters the applications by display name.
    public async filterTreeView(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this._dataProvider.isGraphClientInitialised) {
            this._dataProvider.initialiseGraphClient();
            return;
        }
        // Prompt the user for the filter text.
        const newFilter = await window.showInputBox({
            placeHolder: "Name starts with...",
            prompt: "Filter applications by display name",
            value: this._filterText,
            ignoreFocusOut: true
        });

        // Escape has been hit so we don't want to do anything.
        if ((newFilter === undefined) || (newFilter === '' && newFilter === (this._filterText ?? ""))) {
            return;
        } else if (newFilter === '' && this._filterText !== '') {
            this._filterText = undefined;
            this._filterCommand = undefined;
            await this.populateTreeView(window.setStatusBarMessage("$(loading~spin) Loading application registrations..."));
        } else if (newFilter !== '' && newFilter !== this._filterText) {
            // If the filter text is not empty then set the filter command and filter text.
            this._filterText = newFilter!;
            this._filterCommand = `startsWith(displayName, \'${newFilter}\')`;
            await this.populateTreeView(window.setStatusBarMessage("$(loading~spin) Filtering application registrations..."));
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Application Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Creates a new application registration.
    public async addApp(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this._dataProvider.isGraphClientInitialised) {
            this._dataProvider.initialiseGraphClient();
            return;
        }
        // Add a new application registration and reload the tree view if successful.
        const status = await this._applicationService.add();

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Renames an application registration.
    public async renameApp(item: AppRegItem): Promise<void> {
        // Rename the application registration and reload the tree view if successful.
        const status = await this._applicationService.rename(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Deletes an application registration.
    public async deleteApp(item: AppRegItem): Promise<void> {
        // Delete the application registration and reload the tree view if successful.
        const status = await this._applicationService.delete(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Copies the client Id to the clipboard.
    public copyClientId(item: AppRegItem): void {
        this._applicationService.copyClientId(item);
    }

    // Opens the application registration in the Azure Portal.
    public openAppInPortal(item: AppRegItem): void {
        this._applicationService.openInPortal(item);
    }

    // Opens the application manifest in a new editor window.
    public viewAppManifest(item: AppRegItem): void {
        this._applicationService.viewManifest(item);
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Owner Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new owner to an application registration.
    public async addOwner(item: AppRegItem): Promise<void> {
        // Add a new owner and reload the tree view if successful.
        const status = await this._ownerService.add(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Removes an owner from an application registration.
    public async removeOwner(item: AppRegItem): Promise<void> {
        // Add a new owner and reload the tree view if successful.
        const status = await this._ownerService.remove(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Opens the user in the Azure Portal.
    public openUserInPortal(item: AppRegItem): void {
        this._ownerService.openInPortal(item);
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Sign In Audience Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Edits the application sign in audience.
    public async editAudience(item: AppRegItem): Promise<void> {
        // Edit the sign in audience and reload the tree view if successful.
        const status = await this._signInAudienceService.edit(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Redirect URI Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new redirect URI to an application registration.
    public async addRedirectUri(item: AppRegItem): Promise<void> {
        // Add a new redirect URI and reload the tree view if successful.
        const status = await this._redirectUriService.add(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Deletes a redirect URI.
    public async deleteRedirectUri(item: AppRegItem): Promise<void> {
        // Delete the redirect URI and reload the tree view if successful.
        const status = await this._redirectUriService.delete(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Edits a redirect URI.   
    public async editRedirectUri(item: AppRegItem): Promise<void> {
        // Edit the redirect URI and reload the tree view if successful.
        const status = await this._redirectUriService.edit(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Password Credentials Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a password credential.
    public async addPasswordCredential(item: AppRegItem): Promise<void> {
        // Adds the credential and reload the tree view if successful.
        const status = await this._passwordCredentialService.add(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Deletes a password credential.
    public async deletePasswordCredential(item: AppRegItem): Promise<void> {
        // Delete the credential and reload the tree view if successful.
        const status = await this._passwordCredentialService.delete(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // App Id URI Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Edits the app id URI.   
    public async editAppIdUri(item: AppRegItem): Promise<void> {
        // Edit the app id URI and reload the tree view if successful.
        const status = await this._applicationService.editAppIdUri(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Removes the app id URI.   
    public async removeAppIdUri(item: AppRegItem): Promise<void> {
        // Removes the app id URI and reload the tree view if successful.
        const status = await this._applicationService.removeAppIdUri(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Exposed API Scope Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new exposed api scope to an application registration.
    public async addExposedApiScope(item: AppRegItem): Promise<void> {
        await this._oauth2PermissionScopeService.add(item);
    }

    // Deletes an exposed api scope.
    public async deleteExposedApiScope(item: AppRegItem): Promise<void> {
        await this._oauth2PermissionScopeService.delete(item);
    }

    // Edits an exposed api scope.   
    public async editExposedApiScope(item: AppRegItem): Promise<void> {
        await this._oauth2PermissionScopeService.edit(item);
    }

    // Changes the enabled state of an exposed api scope.   
    public async changeStateExposedApiScope(item: AppRegItem): Promise<void> {
        await this._oauth2PermissionScopeService.changeState(item);
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // App Role Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new app role to an application registration.
    public async addAppRole(item: AppRegItem): Promise<void> {
        await this._appRoleService.add(item);
    }

    // Deletes an app role.
    public async deleteAppRole(item: AppRegItem): Promise<void> {
        await this._appRoleService.delete(item);
    }

    // Edits an app role.   
    public async editAppRole(item: AppRegItem): Promise<void> {
        await this._appRoleService.edit(item);
    }

    // Changes the enabled state of an app role.   
    public async changeStateAppRole(item: AppRegItem): Promise<void> {
        await this._appRoleService.changeState(item);
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Common Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Copies the selected value to the clipboard.
    public copyValue(item: AppRegItem): void {
        env.clipboard.writeText(
            item.contextValue === "COPY"
                || item.contextValue === "WEB-REDIRECT-URI"
                || item.contextValue === "SPA-REDIRECT-URI"
                || item.contextValue === "NATIVE-REDIRECT-URI"
                || item.contextValue === "APPID-URI"
                ? item.value!
                : item.children![0].value!);
    }

    // Handles errors caught by the command handlers.
    private errorHandler(result: ActivityStatus): void {
        // Clear any status bar messages.
        result.statusBarHandle!.dispose();

        // Restore the original icon.
        result.treeViewItem!.iconPath = result.previousIcon;
        this._dataProvider.triggerOnDidChangeTreeData();

        // Display an error message.
        window.showErrorMessage(`An error occurred trying to complete your task: ${result.error!.message}.`);
    }
}