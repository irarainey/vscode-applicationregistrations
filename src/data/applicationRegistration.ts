import * as path from 'path';
import { signInCommandText, view } from '../constants';
import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, Disposable } from 'vscode';
import { Application, KeyCredential, PasswordCredential, User, AppRole, RequiredResourceAccess, PermissionScope } from "@microsoft/microsoft-graph-types";
import { GraphClient } from '../clients/graph';
import { AppRegItem } from '../models/appRegItem';

// This is the application registration data provider for the tree view.
export class AppRegDataProvider implements TreeDataProvider<AppRegItem> {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // Private instance of the tree data
    private _treeData: AppRegItem[] = [];

    // A private instance of the status bar message
    private _statusBarMessage: Disposable | undefined;

    // This is the event that is fired when the tree view is refreshed.
    private _onDidChangeTreeData: EventEmitter<AppRegItem | undefined | null | void> = new EventEmitter<AppRegItem | undefined | null | void>();

    // A private instance of the status bar app count
    private _statusBarAppCount: Disposable | undefined;

    //Defines the event that is fired when the tree view is refreshed.
    public readonly onDidChangeTreeData: Event<AppRegItem | undefined | null | void> = this._onDidChangeTreeData.event;

    // The constructor for the AppRegDataProvider class.
    constructor(graphClient: GraphClient) {
        this._graphClient = graphClient;
        window.registerTreeDataProvider(view, this);
        this.renderTreeView("INITIALISING", undefined, undefined);
        this.initialiseGraphClient(undefined);
    }

    // A public method to initialise the graph client.
    public initialiseGraphClient(statusBar: Disposable | undefined = undefined): void {

        if (statusBar !== undefined) {
            statusBar.dispose();
        }

        this._graphClient.initialiseTreeView = (type: string, statusBarMessage: Disposable | undefined, filter?: string) => { this.renderTreeView(type, statusBarMessage, filter); };
        this._graphClient.initialise();
    }

    // A public get property for the graphClient.
    public get graphClient() {
        return this._graphClient;
    }

    // A public get property for the graphClientInitialised state.
    public get isGraphClientInitialised() {
        return this._graphClient.isGraphClientInitialised;
    }

    // Initialises the tree view data based on the type of data to be displayed.
    public async renderTreeView(type: string, statusBarMessage: Disposable | undefined = undefined, filter?: string): Promise<void> {

        // Clear any existing status bar message
        if (this._statusBarMessage !== undefined) {
            await this._statusBarMessage.dispose();
        }

        this._statusBarMessage = statusBarMessage;

        // Clear the tree data
        this._treeData = [];

        // Add the appropriate tree view item based on the type of data to be displayed.
        switch (type) {
            case "INITIALISING":
                this._treeData.push(new AppRegItem({
                    label: "Initialising extension...",
                    context: "INITIALISING",
                    icon: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
                }));
                this.triggerOnDidChangeTreeData();
                break;
            case "SIGN-IN":
                this._treeData.push(new AppRegItem({
                    label: signInCommandText,
                    context: "SIGN-IN",
                    icon: new ThemeIcon("sign-in", new ThemeColor("editor.foreground")),
                    command: {
                        command: "appRegistrations.signInToAzure",
                        title: signInCommandText,
                    }
                }));
                this.triggerOnDidChangeTreeData();
                break;
            case "APPLICATIONS":
                await this.buildAppRegTree(filter);
                break;
        }
    }

    // Trigger the event to refresh the tree view
    public triggerOnDidChangeTreeData() {
        this._onDidChangeTreeData.fire();
    }

    // Builds the tree view data for the application registrations.
    private async buildAppRegTree(filter?: string): Promise<void> {

        try {
            // Get the application registrations from the graph client.
            const applications = await this.getApplicationList(filter);

            // If we have some applications then add them to the tree view.
            if (applications !== undefined && applications.length > 0) {

                let counter: number = 0;

                // Re-map the application keeping the alphabetical order but changing the insertion order because ES6 only preserves the order of insertion
                // If we don't do this then the tree view will be out of order, despite the items being in the correct order in the array
                let remappedApplications = new Map();
                for (let i = 0; i < applications.length; i++) {
                    remappedApplications.set(applications[i].id, { value: applications[i].displayName, index: i });
                }

                // Now we have the map in the right order we can iterate through it to build our tree view
                remappedApplications.forEach((item, key) => {
                    this._graphClient.getApplicationDetails(key)
                        .then((app: Application) => {
                            // Create the tree view item for the application and it's children
                            this._treeData.push(new AppRegItem({
                                label: app.displayName ? app.displayName : "Application",
                                context: "APPLICATION",
                                icon: path.join(__filename, "..", "..", "..", "resources", "icons", "app.svg"),
                                objectId: app.id!,
                                appId: app.appId!,
                                manifest: app,
                                children: [
                                    // Application (Client) Id
                                    new AppRegItem({
                                        label: "Client Id",
                                        context: "APPID-PARENT",
                                        icon: new ThemeIcon("preview"),
                                        children: [
                                            new AppRegItem({
                                                label: app.appId!,
                                                value: app.appId!,
                                                context: "COPY",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                                            })
                                        ]
                                    }),
                                    // Sign In Audience
                                    new AppRegItem({
                                        label: "Sign In Audience",
                                        context: "AUDIENCE-PARENT",
                                        icon: new ThemeIcon("account"),
                                        objectId: app.id!,
                                        children: [
                                            new AppRegItem({
                                                label: app.signInAudience! === "AzureADMyOrg"
                                                    ? "Single Tenant"
                                                    : app.signInAudience! === "AzureADMultipleOrgs"
                                                        ? "Multi Tenant"
                                                        : "Multi Tenant and Personal Accounts",
                                                context: "AUDIENCE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                objectId: app.id!,
                                            })
                                        ]
                                    }),
                                    // Redirect URIs
                                    new AppRegItem({
                                        label: "Redirect URIs",
                                        context: "REDIRECT-PARENT",
                                        icon: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        children: [
                                            new AppRegItem({
                                                label: "Web",
                                                context: "WEB-REDIRECT",
                                                icon: new ThemeIcon("globe"),
                                                objectId: app.id!,
                                                children: app.web?.redirectUris?.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "SPA",
                                                context: "SPA-REDIRECT",
                                                icon: new ThemeIcon("browser"),
                                                objectId: app.id!,
                                                children: app.spa?.redirectUris?.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "Mobile and Desktop",
                                                context: "NATIVE-REDIRECT",
                                                icon: new ThemeIcon("editor-layout"),
                                                objectId: app.id!,
                                                children: app.publicClient?.redirectUris?.length === 0 ? undefined : []
                                            })
                                        ]
                                    }),
                                    // Credentials
                                    new AppRegItem({
                                        label: "Credentials",
                                        context: "PROPERTY-ARRAY",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                        children: [
                                            new AppRegItem({
                                                label: "Client Secrets",
                                                context: "PASSWORD-CREDENTIALS",
                                                objectId: app.id!,
                                                icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                                children: app.passwordCredentials!.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "Certificates",
                                                context: "CERTIFICATE-CREDENTIALS",
                                                objectId: app.id!,
                                                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                                children: app.keyCredentials!.length === 0 ? undefined : []
                                            })
                                        ]
                                    }),
                                    // API Permissions
                                    new AppRegItem({
                                        label: "API Permissions",
                                        context: "API-PERMISSIONS",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                                        children: app.requiredResourceAccess?.length === 0 ? undefined : []
                                    }),
                                    // Exposed API Permissions
                                    new AppRegItem({
                                        label: "Exposed API Permissions",
                                        context: "EXPOSED-API-PERMISSIONS",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                                        children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : []
                                    }),
                                    // App Roles
                                    new AppRegItem({
                                        label: "App Roles",
                                        context: "APP-ROLES",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                                        children: app.appRoles!.length === 0 ? undefined : []
                                    }),
                                    // Owners
                                    new AppRegItem({
                                        label: "Owners",
                                        context: "OWNERS",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
                                        children: []
                                    }),
                                ]
                            }));

                            counter++;
                        })
                        .then(() => {
                            // Only trigger the redraw event when all applications have been processed
                            if (counter === applications.length) {
                                // Clear any status bar message
                                if (this, this._statusBarMessage !== undefined) {
                                    this._statusBarMessage.dispose();
                                }

                                // Trigger the event to refresh the tree view
                                this.triggerOnDidChangeTreeData();
                            }
                        });
                });
            }
        } catch (error) {
            // if (error.code !== undefined && error.code === "CredentialUnavailableError") {
            //     this._graphClient.isGraphClientInitialised = false;
            //     this._graphClient.initialise();
            // }
            // else {
            //     console.error(error);
            //     window.showErrorMessage(error.message);
            // }
        }
    }

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getTreeItem(element: AppRegItem): TreeItem | Thenable<TreeItem> {
        return element;
    }

    // Returns the UI representation (AppItem) of the element that gets displayed in the view
    public getChildren(element?: AppRegItem | undefined): ProviderResult<AppRegItem[]> {

        // No element selected so return all top level applications to render static elements
        if (element === undefined) {
            return this._treeData;
        }

        // If an element is selected then return the children for that element
        switch (element.contextValue) {
            case "OWNERS":
                // Return the owners for the application
                return this.getApplicationOwners(element);
            case "WEB-REDIRECT":
                // Return the web redirect URIs for the application
                return this.getApplicationRedirectUris(element, "WEB-REDIRECT-URI", this.getParentApplication(element.objectId!).web?.redirectUris!);
            case "SPA-REDIRECT":
                // Return the SPA redirect URIs for the application
                return this.getApplicationRedirectUris(element, "SPA-REDIRECT-URI", this.getParentApplication(element.objectId!).spa?.redirectUris!);
            case "NATIVE-REDIRECT":
                // Return the native redirect URIs for the application
                return this.getApplicationRedirectUris(element, "NATIVE-REDIRECT-URI", this.getParentApplication(element.objectId!).publicClient?.redirectUris!);
            case "PASSWORD-CREDENTIALS":
                // Return the password credentials for the application
                return this.getApplicationPasswordCredentials(element, this.getParentApplication(element.objectId!).passwordCredentials!);
            case "CERTIFICATE-CREDENTIALS":
                // Return the key credentials for the application
                return this.getApplicationKeyCredentials(element, this.getParentApplication(element.objectId!).keyCredentials!);
            case "API-PERMISSIONS":
                // Return the API permissions for the application
                return this.getApplicationApiPermissions(element, this.getParentApplication(element.objectId!).requiredResourceAccess!);
            case "EXPOSED-API-PERMISSIONS":
                // Return the exposed API permissions for the application
                return this.getApplicationExposedApiPermissions(element, this.getParentApplication(element.objectId!).api?.oauth2PermissionScopes!);
            case "APP-ROLES":
                // Return the app roles for the application
                return this.getApplicationAppRoles(element, this.getParentApplication(element.objectId!).appRoles!);
            default:
                // Nothing specific so return the statically defined children
                return element.children;
        }
    }

    // Returns the application registration that is the parent of the given element
    public getParentApplication(objectId: string): Application {
        const app: AppRegItem = this._treeData.filter(item => item.objectId === objectId)[0];
        return app.manifest!;
    }

    // Returns all applications depending on the user setting
    private async getApplicationList(filter?: string): Promise<Application[]> {

        const showAllApplications = workspace.getConfiguration("applicationregistrations").get("showAllApplications") as boolean;

        // If not show all then get only owned applications
        if (showAllApplications === false) {
            const totalApps = await this._graphClient.getApplicationsOwnedCount();
            const response = await this._graphClient.getApplicationNamesOwned(filter);
            this._statusBarAppCount = window.setStatusBarMessage(`$(${filter === undefined ? "filter" : "filter-filled"}) ${response.value.length} of ${totalApps["@odata.count"]} Applications`);
            return response.value;
        } else {
            // Otherwise get all applications
            const totalApps = await this._graphClient.getApplicationsAllCount();
            const response = await this._graphClient.getApplicationNamesAll(filter);
            this._statusBarAppCount = window.setStatusBarMessage(`$(${filter === undefined ? "filter" : "filter-filled"}) ${response.value.length} of ${totalApps["@odata.count"]} Applications`);
            return response.value;
        }
    }

    // Returns the application owners for the given application
    private async getApplicationOwners(element: AppRegItem): Promise<AppRegItem[]> {
        const response = await this._graphClient.getApplicationOwners(element.objectId!);
        const owners: User[] = response.value;
        return owners.map(owner => {
            return new AppRegItem({
                label: owner.displayName!,
                context: "OWNER",
                icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                userId: owner.id!
            });
        });
    }

    // Returns the web redirect URIs for the given application
    private async getApplicationRedirectUris(element: AppRegItem, context: string, redirects: string[]): Promise<AppRegItem[]> {
        return redirects.map(webRedirectUri => {
            return new AppRegItem({
                label: webRedirectUri,
                context: context,
                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                objectId: element.objectId
            });
        });
    }

    // Returns the password credentials for the given application
    private async getApplicationPasswordCredentials(element: AppRegItem, passwords: PasswordCredential[]): Promise<AppRegItem[]> {
        return passwords.map(credential => {
            return new AppRegItem({
                label: credential.displayName!,
                context: "PASSWORD",
                icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                children: [
                    new AppRegItem({
                        label: `Created: ${credential.startDateTime!}`,
                        context: "PASSWORD-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Expires: ${credential.endDateTime!}`,
                        context: "PASSWORD-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    })
                ]
            });
        });
    }

    // Returns the password credentials for the given application
    private async getApplicationKeyCredentials(element: AppRegItem, keys: KeyCredential[]): Promise<AppRegItem[]> {
        return keys.map(credential => {
            return new AppRegItem({
                label: credential.displayName!,
                context: "CERTIFICATE",
                icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                children: [
                    new AppRegItem({
                        label: `Type: ${credential.type!}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Key Identifier: ${credential.customKeyIdentifier!}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Created: ${credential.startDateTime!}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Expires: ${credential.endDateTime!}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    })
                ]
            });
        });
    }

    // Returns the api permissions for the given application
    private async getApplicationApiPermissions(element: AppRegItem, permissions: RequiredResourceAccess[]): Promise<AppRegItem[]> {
        return permissions.map(permission => {
            return new AppRegItem({
                label: permission.resourceAppId!,
                context: "SCOPE",
                icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                children: [
                    new AppRegItem({
                        label: `App Id: ${permission.resourceAppId!}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                     })//,
                    // new AppRegItem({
                    //     label: `Type: ${permission.resourceAccess!}`,
                    //     context: "SCOPE-VALUE",
                    //     icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    // })
                ]
            });
        });
    }
      
    // Returns the exposed api permissions for the given application
    private async getApplicationExposedApiPermissions(element: AppRegItem, permissions: PermissionScope[]): Promise<AppRegItem[]> {
        return permissions.map(permission => {
            return new AppRegItem({
                label: permission.adminConsentDisplayName!,
                context: "SCOPE",
                icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                children: [
                    new AppRegItem({
                        label: `Scope: ${permission.value!}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Description: ${permission.adminConsentDescription!}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Consent: ${permission.type! === "User" ? "Admins and Users" : "Admins Only"}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Enabled: ${permission.isEnabled! ? "Yes" : "No"}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    })
                ]
            });
        });
    }

    // Returns the app roles for the given application
    private async getApplicationAppRoles(element: AppRegItem, roles: AppRole[]): Promise<AppRegItem[]> {
        return roles.map(role => {
            return new AppRegItem({
                label: role.displayName!,
                context: "ROLE",
                icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                value: role.id!,
                children: [
                    new AppRegItem({
                        label: `Value: ${role.value!}`,
                        context: "ROLE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Description: ${role.description!}`,
                        context: "ROLE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Enabled: ${role.isEnabled! ? "Yes" : "No"}`,
                        context: "ROLE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    })
                ]
            });
        });
    }
}
