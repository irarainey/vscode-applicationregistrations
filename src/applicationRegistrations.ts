import * as vscode from 'vscode';
import { execShellCmd } from './utils';
import { portalAppUri, signInAudienceOptions, signInAudienceDocumentation, portalUserUri } from './constants';
import { GraphClient } from './graphClient';
import { User } from "@microsoft/microsoft-graph-types";
import { AppRegDataProvider, AppItem } from './appRegDataProvider';

// This class is responsible for managing the application registrations tree view.
export class ApplicationRegistrations {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // A private string to store the filter command.
    private filterCommand?: string = undefined;

    // A private string to store the filter text.
    private filterText?: string = undefined;

    // A private array to store the subscriptions.
    private subscriptions: vscode.Disposable[] = [];

    // A private boolean to store the authentication state.
    private authenticated: boolean = false;

    // A private function to trigger a change in the authentication state.
    public isUserAuthenticated: (state: boolean | undefined) => void;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider, { subscriptions }: vscode.ExtensionContext) {
        this.graphClient = graphClient;
        this.subscriptions = subscriptions;
        this.dataProvider = dataProvider;
        this.isUserAuthenticated = () => { };
        this.determineAuthenticationState();
    }

    // Determines the authentication state and populates the tree view if authenticated.
    // If not authenticated, the user is prompted to sign in.
    private determineAuthenticationState(): void {
        // Do we think the user is authenticated?
        if (this.authenticated === false) {

            // If not then initialise the GraphClient.
            this.graphClient.initialise();

            // Handle the authentication state change from the graph client.
            this.graphClient.authenticationStateChange = (state: boolean | undefined) => {

                // If the user is authenticated then populate the tree view.
                if (state === true) {
                    this.authenticated = true;
                    this.populateTreeView();
                } else if (state === false) {

                    // If the user is not authenticated then prompt them to sign in.
                    this.dataProvider.initialise("SIGN-IN");

                    // Handle the authentication state change from the sign-in command.
                    this.isUserAuthenticated = (state: boolean | undefined) => {
                        if (state === true) {
                            // If the user has signed in then go back to determine the authentication state.
                            this.determineAuthenticationState();
                        } else if (state === false) {
                            vscode.window.showErrorMessage("Please sign in to Azure CLI.");
                        }
                    };
                }
            };
        } else {
            // If the user is authenticated then just populate the tree view.
            this.populateTreeView();
        }
    }

    // Populates the tree view with the applications.
    public populateTreeView(): void {
        // Get the applications from the GraphClient.
        this.graphClient.getApplicationsAll(this.filterCommand)
            .then((apps) => {
                // Initialise the data provider with the list of applications.
                this.dataProvider.initialise("APPLICATIONS", this.graphClient, apps);
            }).catch(() => {
                // If there is an error then determine the authentication state.
                this.authenticated = false;
                this.determineAuthenticationState();
            });
    };

    // Filters the applications by display name.
    public filterApps(): void {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Prompt the user for the filter text.
        vscode.window.showInputBox({
            placeHolder: "Name starts with...",
            prompt: "Filter applications by display name",
            value: this.filterText
        }).then((newFilter) => {
            // Escape has been hit so we don't want to do anything.
            if ((newFilter === undefined) || (newFilter === '' && newFilter === (this.filterText ?? ""))) {
                return;
            } else if (newFilter === '' && this.filterText !== '') {
                this.filterCommand = undefined;
                this.filterText = undefined;
                this.populateTreeView();
            } else if (newFilter !== '' && newFilter !== this.filterText) {
                // If the filter text is not empty then set the filter command and filter text.
                this.filterText = newFilter!;
                this.filterCommand = `startsWith(displayName, \'${newFilter}\')`;
                this.populateTreeView();
            }
        });
    };

    // Creates a new application registration.
    public addApp(): void {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Prompt the user for the new application name.
        vscode.window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration",
        })
            .then((newName) => {
                // If the new application name is not empty then determine the sign in audience.
                if (newName !== undefined) {
                    // Prompt the user for the sign in audience.
                    vscode.window.showQuickPick(signInAudienceOptions, {
                        placeHolder: "Select the sign in audience...",
                    })
                        .then((audience) => {
                            // If the sign in audience is not undefined then create the application.
                            if (audience !== undefined) {
                                this.graphClient.createApplication({ displayName: newName, signInAudience: this.convertSignInAudience(audience) })
                                    .then(() => {
                                        // If the application is created then populate the tree view.
                                        this.populateTreeView();
                                    }).catch((error) => {
                                        console.error(error);
                                    });
                            }
                        });
                }
            });
    };

    // Renames an application registration.
    public renameApp(app: AppItem): void {
        // Prompt the user for the new application name.
        vscode.window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app.manifest!.displayName!
        })
            .then((newName) => {
                // If the new application name is not empty then update the application.
                if (newName !== undefined) {
                    app.iconPath = new vscode.ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();
                    this.graphClient.updateApplication(app.objectId!, { displayName: newName })
                        .then(() => {
                            // If the application is updated then populate the tree view.
                            this.populateTreeView();
                        }).catch((error) => {
                            console.error(error);
                        });
                }
            });
    };

    // Deletes an application registration.
    public deleteApp(app: AppItem): void {
        // Prompt the user to confirm the deletion.
        vscode.window
            .showInformationMessage(`Do you want to delete the application ${app.label}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    app.iconPath = new vscode.ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();
                    // If the user confirms the deletion then delete the application.
                    this.graphClient.deleteApplication(app.objectId!)
                        .then((response) => {
                            // If the application is deleted then populate the tree view.
                            this.populateTreeView();
                        }).catch((error) => {
                            console.error(error);
                        });
                }
            });
    };

    // Copies the application Id to the clipboard.
    public copyAppId(app: AppItem): void {
        vscode.env.clipboard.writeText(app.appId!);
    };

    // Opens the application registration in the Azure Portal.
    public openAppInPortal(app: AppItem): void {
        vscode.env.openExternal(vscode.Uri.parse(`${portalAppUri}${app.appId}`));
    }

    // Opens the user in the Azure Portal.
    public openUserInPortal(user: AppItem): void {
        vscode.env.openExternal(vscode.Uri.parse(`${portalUserUri}${user.userId}`));
    }

    // Adds a new owner to an application registration.
    public addOwner(item: AppItem): void {
        // Prompt the user for the new owner.
        vscode.window.showInputBox({
            placeHolder: "Enter user name or email address...",
            prompt: "Add new owner to application"
        })
            .then(async (newOwner) => {
                // If the new owner name is not empty then add as an owner.
                if (newOwner !== undefined) {

                    let userList: User[] = [];
                    let identifier: string = "";
                    if (newOwner.indexOf('@') > -1) {
                        // Try to find the user by email.
                        userList = await this.graphClient.findUserByEmail(newOwner);
                        identifier = "email address";
                    } else {
                        // Try to find the user by name.
                        userList = await this.graphClient.findUserByName(newOwner);
                        identifier = "name";
                    }

                    if (userList.length === 0) {
                        // User not found
                        vscode.window.showErrorMessage(`No user with the ${identifier} ${newOwner} was found in your directory.`);
                    } else if (userList.length > 1) {
                        // More than one user found
                        vscode.window.showErrorMessage(`More than one user with the ${identifier} ${newOwner} has been found in your directory.`);
                    } else {
                        // Sweet spot
                        item.iconPath = new vscode.ThemeIcon("loading~spin");
                        this.dataProvider.triggerOnDidChangeTreeData();
                        this.graphClient.addApplicationOwner(item.objectId!, userList[0].id!)
                            .then(() => {
                                // If the application is updated then populate the tree view.
                                this.populateTreeView();
                            }).catch((error) => {
                                console.error(error);
                            });
                    }
                }
            });
    }

    // Removes an owner from an application registration.
    public removeOwner(item: AppItem): void {
        // Prompt the user to confirm the removal.
        vscode.window
            .showInformationMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    // If the user confirms the removal then remove the user.
                    item.iconPath = new vscode.ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();
                    this.graphClient.removeApplicationOwner(item.objectId!, item.userId!)
                        .then(() => {
                            // If the owner is removed then populate the tree view.
                            this.populateTreeView();
                        }).catch((error) => {
                            console.error(error);
                        });
                }
            });
    }

    // Opens the application manifest in a new editor window.
    public async viewAppManifest(app: AppItem): Promise<void> {
        const myProvider = new class implements vscode.TextDocumentContentProvider {
            onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(uri: vscode.Uri): string {
                return JSON.stringify(app.manifest, null, 4);
            }
        };
        this.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider('manifest', myProvider));

        const uri = vscode.Uri.parse('manifest:' + app.label + ".json");
        vscode.workspace.openTextDocument(uri)
            .then(doc => vscode.window.showTextDocument(doc, { preview: false }));
    };

    // Copies the selected value to the clipboard.
    public copyValue(item: AppItem): void {
        vscode.env.clipboard.writeText(item.contextValue === "COPY" ? item.value! : item.children![0].value!);
    };

    // Edits the application sign in audience.
    public editAudience(item: AppItem): void {
        // Prompt the user for the new audience.
        vscode.window.showQuickPick(signInAudienceOptions, {
            placeHolder: "Select the sign in audience...",
        })
            .then((audience) => {
                // If the new audience is not empty then update the application.
                if (audience !== undefined) {
                    // Update the application.
                    if(item.contextValue! === "AUDIENCE-CHILD-EDIT") {
                        item.children![0].iconPath = new vscode.ThemeIcon("loading~spin");
                    } else {
                        item.iconPath = new vscode.ThemeIcon("loading~spin");
                    }
                    this.dataProvider.triggerOnDidChangeTreeData();
                    this.graphClient.updateApplication(item.objectId!, { signInAudience: this.convertSignInAudience(audience) })
                        .then(() => {
                            // If the application is updated then populate the tree view.
                            this.populateTreeView();
                        }).catch(() => {
                            // If the application is not updated then show an error message and a link to the documentation.
                            vscode.window.showErrorMessage(
                                `An error occurred while attempting to change the sign in audience. This is likely because some properties of the application are not supported by the new sign in audience. Please consult the Azure AD documentation for more information at ${signInAudienceDocumentation}.`,
                                ...["OK", "Open Documentation"]
                            )
                                .then((answer) => {
                                    if (answer === "Open Documentation") {
                                        vscode.env.openExternal(vscode.Uri.parse(signInAudienceDocumentation));
                                    }
                                });
                        });
                }
            });
    }

    // Adds a new redirect URI to an application registration.
    public addRedirectUri(item: AppItem): void {
        // Prompt the user for the new redirect URI.
        vscode.window.showInputBox({
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
                item.iconPath = new vscode.ThemeIcon("loading~spin");
                this.dataProvider.triggerOnDidChangeTreeData();
                this.updateRedirectUri(item, existingRedirectUris);
            }
        });
    }

    // Deletes a redirect URI.
    public deleteRedirectUri(uri: AppItem): void {
        // Prompt the user to confirm the deletion.
        vscode.window
            .showInformationMessage(`Do you want to delete the Redirect URI ${uri.label!}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    // Get the parent application so we can read the manifest.
                    const parent = this.dataProvider.getParentApplication(uri.objectId!);
                    let newArray: string[] = [];
                    // Remove the redirect URI from the array.
                    switch (uri.contextValue) {
                        case "WEB-REDIRECT-URI":
                            parent.web!.redirectUris!.splice(parent.web!.redirectUris!.indexOf(uri.label!.toString()), 1);
                            newArray = parent.web!.redirectUris!;
                            break;
                        case "SPA-REDIRECT-URI":
                            parent.spa!.redirectUris!.splice(parent.spa!.redirectUris!.indexOf(uri.label!.toString()), 1);
                            newArray = parent.spa!.redirectUris!;
                            break;
                        case "NATIVE-REDIRECT-URI":
                            parent.publicClient!.redirectUris!.splice(parent.publicClient!.redirectUris!.indexOf(uri.label!.toString()), 1);
                            newArray = parent.publicClient!.redirectUris!;
                            break;
                    }
                    uri.iconPath = new vscode.ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();
                    // Update the application.
                    this.updateRedirectUri(uri, newArray);
                }
            });
    };

    // Edits a redirect URI.   
    public editRedirectUri(uri: AppItem): void {
        // Prompt the user for the new application name.
        vscode.window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: uri.label!.toString()
        })
            .then((updatedUri) => {
                // If the new application name is not empty then update the application.
                if (updatedUri !== undefined && updatedUri !== uri.label!.toString()) {

                    const parent = this.dataProvider.getParentApplication(uri.objectId!);
                    let existingRedirectUris: string[] = [];

                    // Get the existing redirect URIs.
                    switch (uri.contextValue) {
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
                    if (this.validateRedirectUri(updatedUri, uri.contextValue!, existingRedirectUris) === false) {
                        return;
                    }

                    // Remove the old redirect URI and add the new one.
                    existingRedirectUris.splice(existingRedirectUris.indexOf(uri.label!.toString()), 1);
                    existingRedirectUris.push(updatedUri);

                    // Show progress indicator.
                    uri.iconPath = new vscode.ThemeIcon("loading~spin");
                    this.dataProvider.triggerOnDidChangeTreeData();

                    // Update the application.
                    this.updateRedirectUri(uri, existingRedirectUris);
                }
            });
    };

    private updateRedirectUri(item: AppItem, redirectUris: string[]): void {
        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .then(() => {
                    // If the application is updated then populate the tree view.
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .then(() => {
                    // If the application is updated then populate the tree view.
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .then(() => {
                    // If the application is updated then populate the tree view.
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[]): boolean {

        // Check to see if the redirect URI already exists.
        if(existingRedirectUris.includes(uri)) {
            vscode.window.showErrorMessage("The redirect URI specified already exists.");
            return false;
        }

        if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
            // Check the redirect URI starts with https://
            if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
                vscode.window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https:// unless it is using http://localhost.");
                return false;
            }
        }
        else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
            // Check the redirect URI starts with https:// or http:// or customScheme://
            if (uri.includes("://") === false) {
                vscode.window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https, http, or customScheme://.");
                return false;
            }
        }

        // Check the length of the redirect URI.
        if (uri.length > 256) {
            vscode.window.showErrorMessage("The redirect URI is not valid. A redirect URI cannot be longer than 256 characters.");
            return false;
        }

        return true;
    }

    // Invokes the Azure CLI sign-in command.
    public async invokeSignIn(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        vscode.window.showInputBox({
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

    // Converts the audience to the correct format as required for the manifest.
    private convertSignInAudience(audience: string): string {
        return audience === "Single Tenant"
            ? "AzureADMyOrg"
            : audience === "Multiple Tenants"
                ? "AzureADMultipleOrgs"
                : "AzureADandPersonalMicrosoftAccount";
    }
}