import * as path from 'path';
import { signInCommandText, view, appSelectProperties } from '../constants';
import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, Disposable, ConfigurationTarget } from 'vscode';
import { Application, KeyCredential, PasswordCredential, User, AppRole, RequiredResourceAccess, PermissionScope } from "@microsoft/microsoft-graph-types";
import { GraphClient } from '../clients/graph';
import { AppRegItem } from '../models/appRegItem';
import { ActivityResult } from '../interfaces/activityResult';
import { sort } from 'fast-sort';
import { format } from 'date-fns';

// This is the application registration tree data provider for the tree view.
export class AppRegTreeDataProvider implements TreeDataProvider<AppRegItem> {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // Private instance of the tree data
    private _treeData: AppRegItem[] = [];

    // A private instance of the status bar message handle.
    private _statusBarHandle: Disposable | undefined;

    // A private instance of a flag to indicate if the tree view is currently being updated.
    private _isUpdating: boolean = false;

    // This is the event that is fired when the tree view is refreshed.
    private _onDidChangeTreeData: EventEmitter<AppRegItem | undefined | null | void> = new EventEmitter<AppRegItem | undefined | null | void>();

    // A protected instance of the EventEmitter class to handle error events.
    private _onError: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    //Defines the event that is fired when the tree view is refreshed.
    public readonly onDidChangeTreeData: Event<AppRegItem | undefined | null | void> = this._onDidChangeTreeData.event;

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityResult> = this._onError.event;

    // The constructor for the AppRegDataProvider class.
    constructor(graphClient: GraphClient) {
        this._graphClient = graphClient;
        window.registerTreeDataProvider(view, this);
        this.renderTreeView("INITIALISING", undefined, undefined);
        this.initialiseGraphClient(undefined);
    }

    // Trigger the event to indicate an error
    private triggerOnError(item: ActivityResult) {
        this._onError.fire(item);
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

    // Trigger the event to refresh the tree view
    public triggerOnDidChangeTreeData(item?: AppRegItem) {
        this._onDidChangeTreeData.fire(item);
    }

    // Initialises the tree view data based on the type of data to be displayed.
    public async renderTreeView(type: string, statusBarMessage: Disposable | undefined = undefined, filter?: string): Promise<void> {

        // Clear any existing status bar message
        if (this._statusBarHandle !== undefined) {
            await this._statusBarHandle.dispose();
        }

        this._statusBarHandle = statusBarMessage;

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
                this._onDidChangeTreeData.fire(undefined);
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
                this._onDidChangeTreeData.fire(undefined);
                break;
            case "APPLICATIONS":
                await this.populateAppRegTreeData(filter);
                break;
        }
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

            // Get the configuration settings.
            const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
            const showApplicationCountWarning = workspace.getConfiguration("applicationregistrations").get("showApplicationCountWarning") as boolean;

            // Determine if the warning message should be displayed.
            if (showApplicationCountWarning === true) {
                let totalApplicationCount: number = 0;
                const showOwnedApplicationsOnly = workspace.getConfiguration("applicationregistrations").get("showOwnedApplicationsOnly") as boolean;

                // Get the total number of applications in the tenant based on the filter settings.
                if (showOwnedApplicationsOnly) {
                    totalApplicationCount = await this.graphClient.getApplicationCountOwned();
                } else {
                    totalApplicationCount = await this.graphClient.getApplicationCountAll();
                }

                // If the total number of applications is less than 200 and eventual consistency is enabled then display a warning message.
                if (totalApplicationCount <= 200 && useEventualConsistency === true) {
                    window.showWarningMessage(`You have enabled eventual consistency for Graph API calls but only have ${totalApplicationCount} applications in your tenant. You would likely benefit from disabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning")
                        .then((result) => {
                            // If the user selects "Disable Warning" then disable the warning message.
                            if (result === "Disable Warning") {
                                workspace.getConfiguration("applicationregistrations").update("showApplicationCountWarning", false, ConfigurationTarget.Global);
                                // If the user selects "Yes" then disable eventual consistency.
                            } else if (result === "Yes") {
                                workspace.getConfiguration("applicationregistrations").update("useEventualConsistency", false, ConfigurationTarget.Global);
                            }
                        });
                    // If the total number of applications is greater than 200 and eventual consistency is disabled then display a warning message.                        
                } else if (totalApplicationCount > 200 && useEventualConsistency === false) {
                    window.showWarningMessage(`You do not have enabled eventual consistency enabled for Graph API calls and have ${totalApplicationCount} applications in your tenant. You would likely benefit from enabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning")
                        .then((result) => {
                            // If the user selects "Disable Warning" then disable the warning message.
                            if (result === "Disable Warning") {
                                workspace.getConfiguration("applicationregistrations").update("showApplicationCountWarning", false, ConfigurationTarget.Global);
                                // If the user selects "Yes" then enable eventual consistency.
                            } else if (result === "Yes") {
                                workspace.getConfiguration("applicationregistrations").update("useEventualConsistency", true, ConfigurationTarget.Global);
                            }
                        });
                }
            }

            // Get the application registrations from the graph client.
            const applicationList = await this.getApplicationList(filter);

            let allApplications: Application[] = [];

            // If we're not using eventual consistency then we need to sort the application registrations by display name.
            if (useEventualConsistency === false) {
                allApplications = sort(applicationList).asc(a => a.displayName!.toLowerCase().replace(/[`~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, ""));
            } else {
                allApplications = applicationList;
            }

            // Iterate through the application registrations and create the tree view items.
            const unsorted = allApplications.map(async (application, index) => {
                try {
                    // Get the application details.
                    const app = await this._graphClient.getApplicationDetailsPartial(application.id!, appSelectProperties, true);
                    // Create the tree view item.
                    return (new AppRegItem({
                        label: app.displayName!,
                        context: "APPLICATION",
                        icon: path.join(__filename, "..", "..", "..", "resources", "icons", "app.svg"),
                        objectId: app.id!,
                        appId: app.appId!,
                        tooltip: (app.notes !== null ? app.notes! : app.displayName!),
                        order: index,
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
                                children: app.owners!.length === 0 ? undefined : []
                            }),
                        ]
                    }));
                } catch (error: any) {
                    // Need to catch here any attempt to get an application that has been deleted but eventual consistency is still listing it
                    if (error.code !== undefined && error.code === "Request_ResourceNotFound") {
                        console.log("Application no longer exists.");
                    }
                    // Return undefined which we will filter out later
                    return undefined;
                }
            });

            // Sort the applications by name and assign to the class-level array used to render the tree.
            this._treeData = sort(await Promise.all(unsorted.filter(a => a !== undefined)) as AppRegItem[]).asc(async a => a.order);

            // Clear any status bar message
            if (this, this._statusBarHandle !== undefined) {
                this._statusBarHandle.dispose();
            }

            // Trigger the event to refresh the tree view
            this.triggerOnDidChangeTreeData();

            // Set the flag to indicate that the tree is no longer updating
            this._isUpdating = false;

        } catch (error: any) {
            // Set the flag to indicate that the tree is no longer updating
            this._isUpdating = false;

            if (error.code !== undefined && error.code === "CredentialUnavailableError") {
                // Check to see if the user is signed in and if not then prompt them to sign in
                this._graphClient.isGraphClientInitialised = false;
                this._graphClient.initialise();
            }
            else {
                this.triggerOnError({ success: false, statusBarHandle: this._statusBarHandle, error: error });
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
                return this.getApplicationPartial(element.objectId!, "web")
                    .then((app: Application) => {
                        return this.getApplicationRedirectUris(element, "WEB-REDIRECT-URI", app.web!.redirectUris!);
                    });
            case "SPA-REDIRECT":
                // Return the SPA redirect URIs for the application
                return this.getApplicationPartial(element.objectId!, "spa")
                    .then((app: Application) => {
                        return this.getApplicationRedirectUris(element, "SPA-REDIRECT-URI", app.spa!.redirectUris!);
                    });
            case "NATIVE-REDIRECT":
                // Return the native redirect URIs for the application
                return this.getApplicationPartial(element.objectId!, "publicClient")
                    .then((app: Application) => {
                        return this.getApplicationRedirectUris(element, "NATIVE-REDIRECT-URI", app.publicClient!.redirectUris!);
                    });
            case "PASSWORD-CREDENTIALS":
                // Return the password credentials for the application
                return this.getApplicationPartial(element.objectId!, "passwordCredentials")
                    .then((app: Application) => {
                        return this.getApplicationPasswordCredentials(element, app.passwordCredentials!);
                    });
            case "CERTIFICATE-CREDENTIALS":
                // Return the key credentials for the application
                return this.getApplicationPartial(element.objectId!, "keyCredentials")
                    .then((app: Application) => {
                        return this.getApplicationKeyCredentials(element, app.keyCredentials!);
                    });
            case "API-PERMISSIONS":
                // Return the API permissions for the application
                return this.getApplicationPartial(element.objectId!, "requiredResourceAccess")
                    .then((app: Application) => {
                        return this.getApplicationApiPermissions(element, app.requiredResourceAccess!);
                    });
            case "EXPOSED-API-PERMISSIONS":
                // Return the exposed API permissions for the application
                return this.getApplicationPartial(element.objectId!, "api")
                    .then((app: Application) => {
                        return this.getApplicationExposedApiPermissions(element, app.api?.oauth2PermissionScopes!);
                    });
            case "APP-ROLES":
                // Return the app roles for the application
                return this.getApplicationPartial(element.objectId!, "appRoles")
                    .then((app: Application) => {
                        return this.getApplicationAppRoles(element, app.appRoles!);
                    });
            default:
                // Nothing specific so return the statically defined children
                return element.children;
        }
    }

    // Returns the application registration that is the parent of the given element
    public async getApplicationPartial(objectId: string, select: string): Promise<Application> {
        return await this._graphClient.getApplicationDetailsPartial(objectId, select);
    }

    // Returns application list depending on the user setting
    private async getApplicationList(filter?: string): Promise<Application[]> {
        // Get the user setting to determine whether to show all applications or just the ones owned by the user
        const showOwnedApplicationsOnly = workspace.getConfiguration("applicationregistrations").get("showOwnedApplicationsOnly") as boolean;
        return showOwnedApplicationsOnly === true ? await this._graphClient.getApplicationListOwned(filter) : await this._graphClient.getApplicationListAll(filter);
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
    private async getApplicationApiPermissions(element: AppRegItem, permissions: RequiredResourceAccess[]): Promise<AppRegItem[]> {

        // Iterate through each permission and get the service principal app name
        const applicationNames = permissions.map(async (permission) => {
            const response = await this._graphClient.getServicePrincipalByAppId(permission.resourceAppId!);
            return new AppRegItem({
                label: response.displayName!,
                context: "API-PERMISSIONS-APP",
                icon: new ThemeIcon("preview", new ThemeColor("editor.foreground")),
                objectId: element.objectId,
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
    private async getApplicationExposedApiPermissions(element: AppRegItem, scopes: PermissionScope[]): Promise<AppRegItem[]> {
        return scopes.map(scope => {
            const icon = scope.isEnabled! ? "list-tree" : "close";
            return new AppRegItem({
                label: scope.adminConsentDisplayName!,
                context: `SCOPE-${scope.isEnabled! === true ? "ENABLED" : "DISABLED"}`,
                icon: new ThemeIcon(icon, new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                value: scope.id!,
                state: scope.isEnabled!,
                children: [
                    new AppRegItem({
                        label: `Scope: ${scope.value!}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Description: ${scope.adminConsentDescription!}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Consent: ${scope.type! === "User" ? "Admins and Users" : "Admins Only"}`,
                        context: "SCOPE-VALUE",
                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                    }),
                    new AppRegItem({
                        label: `Enabled: ${scope.isEnabled! ? "Yes" : "No"}`,
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
            const icon = role.isEnabled! ? "person" : "close";
            return new AppRegItem({
                label: role.displayName!,
                context: `ROLE-${role.isEnabled! === true ? "ENABLED" : "DISABLED"}`,
                icon: new ThemeIcon(icon, new ThemeColor("editor.foreground")),
                objectId: element.objectId,
                value: role.id!,
                state: role.isEnabled!,
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
                        label: `Allowed: ${role.allowedMemberTypes!.map(type => type === "Application" ? "Applications" : "Users/Groups").join(", ")}`,
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

    // Dispose of the event listener
    public dispose(): void {
        this._onDidChangeTreeData.dispose();
    }
}
