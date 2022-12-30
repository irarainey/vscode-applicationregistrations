import { window, env, Disposable } from 'vscode';
import { AppRegDataProvider } from './data/applicationRegistration';
import { AppRegItem } from './models/appRegItem';
import { ApplicationService } from './services/application';
import { AppRolesService } from './services/appRoles';
import { KeyCredentialsService } from './services/keyCredentials';
import { OAuth2PermissionScopeService } from './services/oauth2PermissionScopes';
import { OwnerService } from './services/owner';
import { PasswordCredentialsService } from './services/passwordCredentials';
import { RedirectUriService } from './services/redirectUris';
import { RequiredResourceAccessService } from './services/requiredResourceAccess';
import { SignInAudienceService } from './services/signInAudience';

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
    private _appRolesService: AppRolesService;
    private _keyCredentialsService: KeyCredentialsService;
    private _oauth2PermissionScopeService: OAuth2PermissionScopeService;
    private _ownerService: OwnerService;
    private _passwordCredentialsService: PasswordCredentialsService;
    private _redirectUriService: RedirectUriService;
    private _requiredResourceAccessService: RequiredResourceAccessService;
    private _signInAudienceService: SignInAudienceService;

    // The constructor for the ApplicationRegistrations class.
    constructor(
        dataProvider: AppRegDataProvider,
        applicationService: ApplicationService,
        appRolesService: AppRolesService,
        keyCredentialsService: KeyCredentialsService,
        oauth2PermissionScopeService: OAuth2PermissionScopeService,
        ownerService: OwnerService,
        passwordCredentialsService: PasswordCredentialsService,
        redirectUriService: RedirectUriService,
        requiredResourceAccessService: RequiredResourceAccessService,
        signInAudienceService: SignInAudienceService) {
        this._dataProvider = dataProvider;
        this._applicationService = applicationService;
        this._appRolesService = appRolesService;
        this._keyCredentialsService = keyCredentialsService;
        this._oauth2PermissionScopeService = oauth2PermissionScopeService;
        this._ownerService = ownerService;
        this._passwordCredentialsService = passwordCredentialsService;
        this._redirectUriService = redirectUriService;
        this._requiredResourceAccessService = requiredResourceAccessService;
        this._signInAudienceService = signInAudienceService;
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Tree View
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Populates the tree view with the applications.
    public async populateTreeView(statusBar: Disposable | undefined = undefined): Promise<void> {
        if (!this._dataProvider.isGraphClientInitialised) {
            this._dataProvider.initialiseGraphClient(statusBar);
            return;
        }
        await this._dataProvider.renderTreeView("APPLICATIONS", statusBar, this._filterCommand);
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
            value: this._filterText
        });

        // Escape has been hit so we don't want to do anything.
        if ((newFilter === undefined) || (newFilter === '' && newFilter === (this._filterText ?? ""))) {
            return;
        } else if (newFilter === '' && this._filterText !== '') {
            this._filterText = undefined;
            this._filterCommand = undefined;
            await this.populateTreeView(window.setStatusBarMessage("$(loading~spin) Loading Application Registrations..."));
        } else if (newFilter !== '' && newFilter !== this._filterText) {
            // If the filter text is not empty then set the filter command and filter text.
            this._filterText = newFilter!;
            this._filterCommand = `startsWith(displayName, \'${newFilter}\')`;
            await this.populateTreeView(window.setStatusBarMessage("$(loading~spin) Filtering Application Registrations..."));
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
        const status = await this._passwordCredentialsService.add(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
    }

    // Deletes a password credential.
    public async deletePasswordCredential(item: AppRegItem): Promise<void> {
        // Delete the credential and reload the tree view if successful.
        const status = await this._passwordCredentialsService.delete(item);

        if (status !== undefined) {
            this.populateTreeView(status);
        }
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
                ? item.value!
                : item.children![0].value!);
    }
}