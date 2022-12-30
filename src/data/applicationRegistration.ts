import * as path from 'path';
import { signInCommandText, view } from '../constants';
import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, Disposable } from 'vscode';
import { Application, KeyCredential, PasswordCredential, User, AppRole, RequiredResourceAccess, PermissionScope } from "@microsoft/microsoft-graph-types";
import { GraphClient } from '../clients/graph';
import { AppRegItem } from '../models/appRegItem';
import { sort } from 'fast-sort';
import { format } from 'date-fns';

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

    // A private instance of a flag to indicate if the tree view is currently being updated.
    private _isUpdating: boolean = false;

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
                await this.populateAppRegTreeData(filter);
                break;
        }
    }

    // Trigger the event to refresh the tree view
    public triggerOnDidChangeTreeData() {
        this._onDidChangeTreeData.fire();
    }

    // Populates the tree view data for the application registrations.
    private async populateAppRegTreeData(filter?: string): Promise<void> {

        // If the tree view is already being updated then return.
        if (this._isUpdating) {
            return;
        }

        try {
            // Set the flag to indicate that the tree view is being updated.
            this._isUpdating = true;

            // Get the application registrations from the graph client.
            // This method uses eventual consistency to allow us to order the applications by display name before
            // the filter is applied and the maximum number of applications is returned. But this is problematic
            // for the rest of the properties, so we're only going to get the ids and then get the details for each
            // in a separate call without using eventual consistency.
            const applications = await this.getApplicationList(filter);

            // If we have some applications then add them to the tree view.
            if (applications !== undefined && applications.length > 0) {

                // Create an empty array to hold the tree view items.
                let unsorted: AppRegItem[] = [];

                // Loop through the applications and get the details for each one.
                for (let item of applications) {

                    // Get the details for the application without using eventual consistency.
                    this._graphClient.getApplicationDetails(item.id!)
                        .then((app: Application) => {
                            // Populate an array with tree view items for the application and it's static children
                            unsorted.push(new AppRegItem({
                                label: app.displayName!,
                                context: "APPLICATION",
                                icon: path.join(__filename, "..", "..", "..", "resources", "icons", "app.svg"),
                                objectId: app.id!,
                                appId: app.appId!,
                                tooltip: (app.notes !== null ? app.notes! : app.displayName!),
                                manifest: app,
                                children: [
                                    // Application (Client) Id
                                    new AppRegItem({
                                        label: "Client Id",
                                        context: "APPID-PARENT",
                                        icon: new ThemeIcon("preview"),
                                        tooltip: "The Application (Client) Id is used to identify the application to Azure AD.",
                                        children: [
                                            new AppRegItem({
                                                label: app.appId!,
                                                value: app.appId!,
                                                context: "COPY",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                tooltip: "The Application (Client) Id is used to identify the application to Azure AD.",
                                            })
                                        ]
                                    }),
                                    // Application ID URI
                                    new AppRegItem({
                                        label: "Application Id URI",
                                        context: "APPID-URI-PARENT",
                                        objectId: app.id!,
                                        appId: app.appId!,
                                        value: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
                                        icon: new ThemeIcon("globe"),
                                        tooltip: "The Application Id URI is a globally unique URI used to identify this web API. It is the prefix for scopes and in access tokens, it is the value of the audience claim. Also referred to as an identifier URI.",
                                        children: [
                                            new AppRegItem({
                                                label: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
                                                value: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
                                                appId: app.appId!,
                                                objectId: app.id!,
                                                context: "APPID-URI",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                tooltip: "The Application Id URI, this is set when an application is used as a resource app. The URI acts as the prefix for the scopes you'll reference in your API's code, and must be globally unique.",
                                            })
                                        ]
                                    }),
                                    // Sign In Audience
                                    new AppRegItem({
                                        label: "Sign In Audience",
                                        context: "AUDIENCE-PARENT",
                                        icon: new ThemeIcon("account"),
                                        objectId: app.id!,
                                        tooltip: "The Sign In Audience determines whether the application can be used by accounts in the same Azure AD tenant or accounts in any Azure AD tenant.",
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
                                                tooltip: "The Sign In Audience determines whether the application can be used by accounts in the same Azure AD tenant or accounts in any Azure AD tenant.",
                                            })
                                        ]
                                    }),
                                    // Redirect URIs
                                    new AppRegItem({
                                        label: "Redirect URIs",
                                        context: "REDIRECT-PARENT",
                                        icon: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        tooltip: "The Redirect URIs are the endpoints where Azure AD will return responses containing tokens or errors.",
                                        children: [
                                            new AppRegItem({
                                                label: "Web",
                                                context: "WEB-REDIRECT",
                                                icon: new ThemeIcon("globe"),
                                                objectId: app.id!,
                                                tooltip: "Redirect URIs for web applications.",
                                                children: app.web?.redirectUris?.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "SPA",
                                                context: "SPA-REDIRECT",
                                                icon: new ThemeIcon("browser"),
                                                objectId: app.id!,
                                                tooltip: "Redirect URIs for single page applications.",
                                                children: app.spa?.redirectUris?.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "Mobile and Desktop",
                                                context: "NATIVE-REDIRECT",
                                                icon: new ThemeIcon("editor-layout"),
                                                objectId: app.id!,
                                                tooltip: "Redirect URIs for mobile and desktop applications.",
                                                children: app.publicClient?.redirectUris?.length === 0 ? undefined : []
                                            })
                                        ]
                                    }),
                                    // Credentials
                                    new AppRegItem({
                                        label: "Credentials",
                                        context: "PROPERTY-ARRAY",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("shield", new ThemeColor("editor.foreground")),
                                        tooltip: "Credentials enable confidential applications to identify themselves to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
                                        children: [
                                            new AppRegItem({
                                                label: "Client Secrets",
                                                context: "PASSWORD-CREDENTIALS",
                                                objectId: app.id!,
                                                icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                                tooltip: "Client secrets are used to authenticate confidential applications to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
                                                children: app.passwordCredentials!.length === 0 ? undefined : []
                                            }),
                                            new AppRegItem({
                                                label: "Certificates",
                                                context: "CERTIFICATE-CREDENTIALS",
                                                objectId: app.id!,
                                                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                                tooltip: "Certificates are used to authenticate confidential applications to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
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
                                        tooltip: "API permissions define the access that an application has to an API.",
                                        children: app.requiredResourceAccess?.length === 0 ? undefined : []
                                    }),
                                    // Exposed API Permissions
                                    new AppRegItem({
                                        label: "Exposed API Permissions",
                                        context: "EXPOSED-API-PERMISSIONS",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                                        tooltip: "Exposed API permissions define custom scopes to restrict access to data and functionality protected by this API.",
                                        children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : []
                                    }),
                                    // App Roles
                                    new AppRegItem({
                                        label: "App Roles",
                                        context: "APP-ROLES",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                                        tooltip: "App roles define custom roles that can be assigned to users and groups, or assigned as application-only scopes to a client application.",
                                        children: app.appRoles!.length === 0 ? undefined : []
                                    }),
                                    // Owners
                                    new AppRegItem({
                                        label: "Owners",
                                        context: "OWNERS",
                                        objectId: app.id!,
                                        icon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
                                        tooltip: "Owners are users who can manage the application.",
                                        children: []
                                    }),
                                ]
                            }));
                        })
                        .then(() => {
                            // Only trigger the redraw event when all applications have been processed
                            if (unsorted.length === applications.length) {

                                // Sort the applications by name and assign to the class-level array used to render the tree.
                                // This is required to ensure the order is maintained.
                                this._treeData = sort(unsorted).asc(a => a.label);

                                // Clear any status bar message
                                if (this, this._statusBarMessage !== undefined) {
                                    this._statusBarMessage.dispose();
                                }

                                // Trigger the event to refresh the tree view
                                this.triggerOnDidChangeTreeData();

                                // Set the flag to indicate that the tree is no longer updating
                                this._isUpdating = false;
                            }
                        });
                }
            }
        } catch (error: any) {
            // Set the flag to indicate that the tree is no longer updating
            this._isUpdating = false;

            // Check to see if the user is signed in and if not then prompt them to sign in
            if (error.code !== undefined && error.code === "CredentialUnavailableError") {
                this._graphClient.isGraphClientInitialised = false;
                this._graphClient.initialise();
            }
            else {
                // Otherwise just show an error message
                console.error(error);
                window.showErrorMessage(error.message);
            }
        }
    }

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
                return this.getApplicationApiPermissions(this.getParentApplication(element.objectId!).requiredResourceAccess!);
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
                userId: owner.id!,
                children: [
                    new AppRegItem({
                        label: owner.mail!,
                        context: "COPY",
                        value: owner.mail!,
                        icon: new ThemeIcon("mail", new ThemeColor("editor.foreground")),
                        tooltip: "Email address of user"
                    }),
                    new AppRegItem({
                        label: owner.userPrincipalName!,
                        context: "COPY",
                        value: owner.userPrincipalName!,
                        icon: new ThemeIcon("account", new ThemeColor("editor.foreground")),
                        tooltip: "User principal name of user"
                    })
                ]
            });
        });
    }

    // Returns the redirect URIs for the given application
    private async getApplicationRedirectUris(element: AppRegItem, context: string, redirects: string[]): Promise<AppRegItem[]> {
        return redirects.map(redirectUri => {
            return new AppRegItem({
                label: redirectUri,
                value: redirectUri,
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
                value: credential.keyId!,
                objectId: element.objectId,
                tooltip: "A secret string that the application uses to prove its identity when requesting a token. Also can be referred to as application password.",
                children: [
                    new AppRegItem({
                        label: `Value: ${credential.hint!}******************`,
                        context: "PASSWORD-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                        tooltip: "Client secret values cannot be viewed or accessed, except for immediately after creation."
                    }),
                    new AppRegItem({
                        label: `Created: ${format(new Date(credential.startDateTime!), 'yyyy-MM-dd')}`,
                        context: "PASSWORD-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Expires: ${format(new Date(credential.endDateTime!), 'yyyy-MM-dd')}`,
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
                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                tooltip: "A certificate that the application uses to prove its identity when requesting a token. Also can be referred to as application certificate or public key.",
                children: [
                    new AppRegItem({
                        label: `Type: ${credential.type!}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Key Identifier: ${credential.customKeyIdentifier!}`,
                        context: "COPY",
                        value: credential.customKeyIdentifier!,
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Created: ${format(new Date(credential.startDateTime!), 'yyyy-MM-dd')}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Expires: ${format(new Date(credential.endDateTime!), 'yyyy-MM-dd')}`,
                        context: "CERTIFICATE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    })
                ]
            });
        });
    }

    // Returns the api permissions for the given application
    private async getApplicationApiPermissions(permissions: RequiredResourceAccess[]): Promise<AppRegItem[]> {

        // Iterate through each permission and get the service principal app name
        const applicationNames = permissions.map(async (permission) => {
            const response = await this._graphClient.getServicePrincipalByAppId(permission.resourceAppId!);
            return new AppRegItem({
                label: response.displayName!,
                context: "API-PERMISSIONS-APP",
                icon: new ThemeIcon("preview", new ThemeColor("editor.foreground")),
                value: permission.resourceAppId,
                children: permission.resourceAccess!.map(resourceAccess => {

                    let scopeLabel = "";
                    let tooltip = undefined;
                    if (resourceAccess.type === "Scope") {
                        scopeLabel = `Delegated: ${response.oauth2PermissionScopes!.find(scope => scope.id === resourceAccess.id)!.value!}`;
                        tooltip = response.oauth2PermissionScopes!.find(scope => scope.id === resourceAccess.id)!.adminConsentDescription;
                    } else {
                        scopeLabel = `Application: ${response.appRoles!.find(scope => scope.id === resourceAccess.id)!.value!}`;
                        tooltip = response.appRoles!.find(scope => scope.id === resourceAccess.id)!.description;
                    }

                    return new AppRegItem({
                        label: scopeLabel,
                        context: "API-PERMISSIONS-SCOPE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                        tooltip: tooltip!
                    });
                })
            });
        });

        return Promise.all(applicationNames);
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
