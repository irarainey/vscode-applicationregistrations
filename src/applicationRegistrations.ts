import { ExtensionContext, window, ThemeIcon, env, Uri, Disposable } from 'vscode';
import { execShellCmd } from './utils/shellUtils';
import { GraphClient } from './clients/graph';
import { AppRegDataProvider } from './dataProviders/applicationRegistration';
import { AppRegItem } from './models/appRegItem';
import { ApplicationService } from './services/application';
import { OwnerService } from './services/owner';
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
    private ownerService: OwnerService;
    private signInAudienceService: SignInAudienceService;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider, context: ExtensionContext) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
        this.applicationService = new ApplicationService(graphClient, dataProvider, context);
        this.ownerService = new OwnerService(graphClient, dataProvider);
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
    public addRedirectUri(item: AppRegItem): void {
        // Prompt the user for the new redirect URI.
        window.showInputBox({
            placeHolder: "Enter redirect URI...",
            prompt: "Add a new redirect URI to the application"
        }).then((redirectUri) => {
            // If the redirect URI is not empty then add it to the application.
            if (redirectUri !== undefined && redirectUri.length > 0) {

                let existingRedirectUris: string[] = item.children!.map((child) => {
                    return child.label!.toString();
                });

                // Validate the redirect URI.
                if (this.validateRedirectUri(redirectUri, item.contextValue!, existingRedirectUris) === false) {
                    return;
                }

                existingRedirectUris.push(redirectUri);
                item.iconPath = new ThemeIcon("loading~spin");
                this.dataProvider.triggerOnDidChangeTreeData();
                this.updateRedirectUri(item, existingRedirectUris);
            }
        });
    }

    // Deletes a redirect URI.
    public deleteRedirectUri(item: AppRegItem): void {
        // Prompt the user to confirm the deletion.
        window
            .showInformationMessage(`Do you want to delete the Redirect URI ${item.label!}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    // Get the parent application so we can read the manifest.
                    const parent = this.dataProvider.getParentApplication(item.objectId!);
                    let newArray: string[] = [];
                    // Remove the redirect URI from the array.
                    switch (item.contextValue) {
                        case "WEB-REDIRECT-URI":
                            parent.web!.redirectUris!.splice(parent.web!.redirectUris!.indexOf(item.label!.toString()), 1);
                            newArray = parent.web!.redirectUris!;
                            break;
                        case "SPA-REDIRECT-URI":
                            parent.spa!.redirectUris!.splice(parent.spa!.redirectUris!.indexOf(item.label!.toString()), 1);
                            newArray = parent.spa!.redirectUris!;
                            break;
                        case "NATIVE-REDIRECT-URI":
                            parent.publicClient!.redirectUris!.splice(parent.publicClient!.redirectUris!.indexOf(item.label!.toString()), 1);
                            newArray = parent.publicClient!.redirectUris!;
                            break;
                    }
                    item.iconPath = new ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();
                    // Update the application.
                    this.updateRedirectUri(item, newArray);
                }
            });
    };

    // Edits a redirect URI.   
    public editRedirectUri(item: AppRegItem): void {
        // Prompt the user for the new application name.
        window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: item.label!.toString()
        })
            .then((updatedUri) => {
                // If the new application name is not empty then update the application.
                if (updatedUri !== undefined && updatedUri !== item.label!.toString()) {

                    const parent = this.dataProvider.getParentApplication(item.objectId!);
                    let existingRedirectUris: string[] = [];

                    // Get the existing redirect URIs.
                    switch (item.contextValue) {
                        case "WEB-REDIRECT-URI":
                            existingRedirectUris = parent.web!.redirectUris!;
                            break;
                        case "SPA-REDIRECT-URI":
                            existingRedirectUris = parent.spa!.redirectUris!;
                            break;
                        case "NATIVE-REDIRECT-URI":
                            existingRedirectUris = parent.publicClient!.redirectUris!;
                            break;
                    }

                    // Validate the edited redirect URI.
                    if (this.validateRedirectUri(updatedUri, item.contextValue!, existingRedirectUris) === false) {
                        return;
                    }

                    // Remove the old redirect URI and add the new one.
                    existingRedirectUris.splice(existingRedirectUris.indexOf(item.label!.toString()), 1);
                    existingRedirectUris.push(updatedUri);

                    // Show progress indicator.
                    item.iconPath = new ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();

                    // Update the application.
                    this.updateRedirectUri(item, existingRedirectUris);
                }
            });
    };

    private async updateRedirectUri(item: AppRegItem, redirectUris: string[]): Promise<void> {
        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .then(async () => {
                    // If the application is updated then populate the tree view.
                    await this.loadTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .then(async () => {
                    // If the application is updated then populate the tree view.
                    await this.loadTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .then(async () => {
                    // If the application is updated then populate the tree view.
                    await this.loadTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[]): boolean {

        // Check to see if the redirect URI already exists.
        if (existingRedirectUris.includes(uri)) {
            window.showErrorMessage("The redirect URI specified already exists.");
            return false;
        }

        if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
            // Check the redirect URI starts with https://
            if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
                window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https:// unless it is using http://localhost.");
                return false;
            }
        }
        else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
            // Check the redirect URI starts with https:// or http:// or customScheme://
            if (uri.includes("://") === false) {
                window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https, http, or customScheme://.");
                return false;
            }
        }

        // Check the length of the redirect URI.
        if (uri.length > 256) {
            window.showErrorMessage("The redirect URI is not valid. A redirect URI cannot be longer than 256 characters.");
            return false;
        }

        return true;
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Common Commands
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Copies the selected value to the clipboard.
    public copyValue(item: AppRegItem): void {
        env.clipboard.writeText(item.contextValue === "COPY" ? item.value! : item.children![0].value!);
    };
}