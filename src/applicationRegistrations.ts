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
    private filterText: string = '';

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
                    this.dataProvider.initialise("SIGNIN");

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
        // Initialise the data provider with the loading state.
        this.dataProvider.initialise("LOADING");

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
    public async filterApps(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Prompt the user for the filter text.
        const filterText = await vscode.window.showInputBox({
            placeHolder: "Name starts with...",
            prompt: "Filter applications by display name",
            value: this.filterText
        });

        // If the filter text is empty or undefined then clear the filter.
        if (filterText === '' || filterText === undefined) {
            this.filterCommand = undefined;
            this.filterText = '';
            this.populateTreeView();
        } else {
            // If the filter text is not empty then set the filter command and filter text.
            this.filterText = filterText;
            this.filterCommand = `startsWith(displayName, \'${filterText}\')`;
            this.populateTreeView();
        }
    };

    // Creates a new application registration.
    public async addApp(): Promise<void> {
        // If the user is not authenticated then we don't want to do anything        
        if (!this.authenticated) {
            return;
        }

        // Prompt the user for the new application name.
        const newName = await vscode.window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration",
        });

        // If the new application name is not empty then determine the sign in audience.
        if (newName !== undefined) {
            // Prompt the user for the sign in audience.
            const audience = await vscode.window.showQuickPick(signInAudienceOptions, {
                placeHolder: "Select the sign in audience...",
            });

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
        }
    };

    // Renames an application registration.
    public async renameApp(app: AppItem): Promise<void> {
        // Prompt the user for the new application name.
        const newName = await vscode.window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app.manifest!.displayName!
        });

        // If the new application name is not empty then update the application.
        if (newName !== undefined) {
            this.graphClient.updateApplication(app.objectId!, { displayName: newName })
                .then(() => {
                    // If the application is updated then populate the tree view.
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
    };

    // Deletes an application registration.
    public deleteApp(app: AppItem): void {
        // Prompt the user to confirm the deletion.
        vscode.window
            .showInformationMessage(`Do you want to delete the application ${app.label}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
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
    public async addOwner(item: AppItem): Promise<void> {
        // Prompt the user for the new owner.
        const newOwner = await vscode.window.showInputBox({
            placeHolder: "Enter user name or email address...",
            prompt: "Add new owner to application"
        });

        // If the new owner name is not empty then add as an owner.
        if (newOwner !== undefined) {

            let userList: User[] = [];
            let identifier: string = "";
            if(newOwner.indexOf('@') > -1) {
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
                this.graphClient.addApplicationOwner(item.objectId!, userList[0].id!)
                    .then(() => {
                        // If the application is updated then populate the tree view.
                        this.populateTreeView();
                    }).catch((error) => {
                        console.error(error);
                    });
            }
        }
    }

    // Removes an owner from an application registration.
    public removeOwner(item: AppItem): void {
        // Prompt the user to confirm the removal.
        vscode.window
            .showInformationMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    // If the user confirms the removal then remove the user.
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
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.window.showTextDocument(doc, { preview: false });
    };

    // Copies the selected value to the clipboard.
    public copyValue(item: AppItem): void {
        vscode.env.clipboard.writeText(item.contextValue === "COPY" ? item.value! : item.children![0].value!);
    };

    // Edits the application sign in audience.
    public async editAudience(item: AppItem): Promise<void> {
        // Prompt the user for the new audience.
        const audience = await vscode.window.showQuickPick(signInAudienceOptions, {
            placeHolder: "Select the sign in audience...",
        });

        // If the new audience is not empty then update the application.
        if (audience !== undefined) {
            // Convert the audience to the correct format.
            // Update the application.
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
    }

    // Invokes the Azure CLI sign-in command.
    public async invokeSignIn(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        const tenant = await vscode.window.showInputBox({
            placeHolder: "Tenant name or Id...",
            prompt: "Enter the tenant name or Id, or leave blank for the default tenant",
        });

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