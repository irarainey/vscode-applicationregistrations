import { ExtensionContext, window, ThemeIcon, env, Uri, Disposable } from 'vscode';
import { execShellCmd } from './utils/shellUtils';
import { GraphClient } from './clients/graph';
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

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // A private string to store the filter command.
    private filterCommand?: string = undefined;

    // A private string to store the filter text.
    private filterText?: string = undefined;

    // A private boolean to store the authentication state.
    private authenticated: boolean = false;

    // A private function to trigger a change in the authentication state.
    public isUserAuthenticated: (state: boolean | undefined) => void;

    // Private instances of our services
    private applicationService: ApplicationService;
    private appRolesService: AppRolesService;
    private keyCredentialsService: KeyCredentialsService;
    private oauth2PermissionScopeService: OAuth2PermissionScopeService;
    private ownerService: OwnerService;
    private passwordCredentialsService: PasswordCredentialsService;
    private redirectUriService: RedirectUriService;
    private requiredResourceAccessService: RequiredResourceAccessService;
    private signInAudienceService: SignInAudienceService;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider, context: ExtensionContext) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
        this.applicationService = new ApplicationService(graphClient, dataProvider, context);
        this.appRolesService = new AppRolesService(graphClient, dataProvider);
        this.keyCredentialsService = new KeyCredentialsService(graphClient, dataProvider);
        this.oauth2PermissionScopeService = new OAuth2PermissionScopeService(graphClient, dataProvider);
        this.ownerService = new OwnerService(graphClient, dataProvider);
        this.passwordCredentialsService = new PasswordCredentialsService(graphClient, dataProvider);
        this.redirectUriService = new RedirectUriService(graphClient, dataProvider);
        this.requiredResourceAccessService = new RequiredResourceAccessService(graphClient, dataProvider);
        this.signInAudienceService = new SignInAudienceService(graphClient, dataProvider);

        this.isUserAuthenticated = () => { };
        this.authenticate();
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Authentication
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Determines the authentication state and populates the tree view if authenticated.
    // If not authenticated, the user is prompted to sign in.
    private async authenticate(): Promise<void> {
        // Do we think the user is authenticated?
        if (this.authenticated === false) {
            // If not then initialise the GraphClient.
            this.graphClient.initialise();

            // Handle the authentication state change from the graph client.
            this.graphClient.authenticationStateChange = (state: boolean | undefined) => {

                // If the user is authenticated then populate the tree view.
                if (state === true) {
                    this.authenticated = true;
                    this.loadTreeView(window.setStatusBarMessage("$(loading~spin) Loading Application Registrations..."));
                } else if (state === false) {

                    // If the user is not authenticated then prompt them to sign in.
                    const status = window.setStatusBarMessage("$(loading~spin) Waiting for Azure CLI sign in...");
                    this.dataProvider.initialise("SIGN-IN", status);

                    // Handle the authentication state change from the sign-in command.
                    this.isUserAuthenticated = (state: boolean | undefined) => {
                        if (state === true) {
                            // If the user has signed in then go back to determine the authentication state.
                            this.authenticate();
                        } else if (state === false) {
                            window.showErrorMessage("Please sign in to Azure CLI.");
                        }
                    };
                }
            };
        } else {
            // If the user is authenticated then just populate the tree view.
            this.loadTreeView(window.setStatusBarMessage("$(loading~spin) Loading Application Registrations..."));
        }
    }

    // Invokes the Azure CLI sign-in command.
    public async cliSignIn(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        window.showInputBox({
            placeHolder: "Tenant name or Id...",
            prompt: "Enter the tenant name or Id, or leave blank for the default tenant",
        })
            .then((tenant) => {
                // If the tenant is undefined then we don't want to do anything because they pressed cancel.
                if (tenant === undefined) {
                    return;
                }

                // Build the command to invoke the Azure CLI sign-in command.
                let command = "az login";
                if (tenant.length > 0) {
                    command += ` --tenant ${tenant}`;
                }

                // Execute the command.
                execShellCmd(command)
                    .then(() => {
                        // If the sign-in is successful then trigger the authentication state change.
                        this.isUserAuthenticated(true);
                    }).catch(() => {
                        this.isUserAuthenticated(false);
                    });
            });
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Tree View
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Populates the tree view with the applications.
    public async loadTreeView(statusBar: Disposable | undefined = undefined): Promise<void> {
        // Get the applications from the GraphClient.
        await this.graphClient.getApplicationsAll(this.filterCommand)
            .then((apps) => {
                // Initialise the data provider with the list of applications.
                this.dataProvider.initialise("APPLICATIONS", statusBar, this.graphClient, apps);
            }).catch(() => {
                // If there is an error then determine the authentication state.
                this.authenticated = false;
                this.authenticate();
            });
    };

    // Filters the applications by display name.
    public async filterTreeView(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Prompt the user for the filter text.
        const newFilter = await window.showInputBox({
            placeHolder: "Name starts with...",
            prompt: "Filter applications by display name",
            value: this.filterText
        });

        // Escape has been hit so we don't want to do anything.
        if ((newFilter === undefined) || (newFilter === '' && newFilter === (this.filterText ?? ""))) {
            return;
        } else if (newFilter === '' && this.filterText !== '') {
            this.filterCommand = undefined;
            this.filterText = undefined;
            await this.loadTreeView(window.setStatusBarMessage("$(loading~spin) Loading Application Registrations..."));
        } else if (newFilter !== '' && newFilter !== this.filterText) {
            // If the filter text is not empty then set the filter command and filter text.
            this.filterText = newFilter!;
            this.filterCommand = `startsWith(displayName, \'${newFilter}\')`;
            await this.loadTreeView(window.setStatusBarMessage("$(loading~spin) Filtering Application Registrations..."));
        }
    };

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Application Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Creates a new application registration.
    public async addApp(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Add a new application registration and reload the tree view if successful.
        await this.applicationService.add()
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    };

    // Renames an application registration.
    public async renameApp(item: AppRegItem): Promise<void> {
        // Rename the application registration and reload the tree view if successful.
        await this.applicationService.rename(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    }

    // Deletes an application registration.
    public async deleteApp(item: AppRegItem): Promise<void> {
        // Delete the application registration and reload the tree view if successful.
        await this.applicationService.delete(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    };

    // Copies the application Id to the clipboard.
    public copyAppId(item: AppRegItem): void {
        this.applicationService.copyId(item);
    };

    // Opens the application registration in the Azure Portal.
    public openAppInPortal(item: AppRegItem): void {
        this.applicationService.openInPortal(item);
    }

    // Opens the application manifest in a new editor window.
    public viewAppManifest(item: AppRegItem): void {
        this.applicationService.viewManifest(item);
    };

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Owner Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new owner to an application registration.
    public async addOwner(item: AppRegItem): Promise<void> {
        // Add a new owner and reload the tree view if successful.
        await this.ownerService.add(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    }

    // Removes an owner from an application registration.
    public async removeOwner(item: AppRegItem): Promise<void> {
        // Add a new owner and reload the tree view if successful.
        await this.ownerService.remove(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    }

    // Opens the user in the Azure Portal.
    public openUserInPortal(item: AppRegItem): void {
        this.ownerService.openInPortal(item);
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Sign In Audience Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Edits the application sign in audience.
    public async editAudience(item: AppRegItem): Promise<void> {
        // Edit the sign in audience and reload the tree view if successful.
        await this.signInAudienceService.edit(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Redirect URI Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Adds a new redirect URI to an application registration.
    public async addRedirectUri(item: AppRegItem): Promise<void> {
        // Add a new redirect URI and reload the tree view if successful.
        await this.redirectUriService.add(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    }

    // Deletes a redirect URI.
    public async deleteRedirectUri(item: AppRegItem): Promise<void> {
        // Delete the redirect URI and reload the tree view if successful.
        await this.redirectUriService.delete(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    };

    // Edits a redirect URI.   
    public async editRedirectUri(item: AppRegItem): Promise<void> {
        // Edit the redirect URI and reload the tree view if successful.
        await this.redirectUriService.edit(item)
            .then((status) => {
                if (status !== undefined) {
                    this.loadTreeView(status);
                }
            });
    };

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Common Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Copies the selected value to the clipboard.
    public copyValue(item: AppRegItem): void {
        env.clipboard.writeText(item.contextValue === "COPY" ? item.value! : item.children![0].value!);
    };
}