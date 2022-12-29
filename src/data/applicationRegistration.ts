import * as path from 'path';
import { signInCommandText, view } from '../constants';
import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, Disposable } from 'vscode';
import { Application, User } from "@microsoft/microsoft-graph-types";
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
        // Get the application registrations from the graph client.
        await this.getApplications(filter)
            .then((apps) => {
                // Iterate through the applications and create the tree data
                apps!.forEach(app => {
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
                                        icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                        children: app.keyCredentials!.length === 0 ? undefined : []
                                    })
                                ]
                            }),
                            // API Permissions
                            new AppRegItem({
                                label: "API Permissions",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                                children: app.requiredResourceAccess?.length === 0 ? undefined : []
                            }),
                            // Exposed API Permissions
                            new AppRegItem({
                                label: "Exposed API Permissions",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                                children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : []
                            }),
                            // App Roles
                            new AppRegItem({
                                label: "App Roles",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                                children: app.appRoles!.length === 0 ? undefined : []
                            }),
                            // Owners
                            new AppRegItem({
                                label: "Owners",
                                context: "OWNERS",
                                icon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                children: []
                            }),
                        ]
                    }));
                });

            })
            .then(() => {
                // Clear any status bar message
                if (this, this._statusBarMessage !== undefined) {
                    this._statusBarMessage.dispose();
                }

                // Trigger the event to refresh the tree view
                this.triggerOnDidChangeTreeData();
            })
            .catch((error) => {
                if (error.code !== undefined && error.code === "CredentialUnavailableError") {
                    this._graphClient.isGraphClientInitialised = false;
                    this._graphClient.initialise();
                }
                else {
                    console.error(error);
                    window.showErrorMessage(error.message);
                }
            });
    }

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getTreeItem(element: AppRegItem): TreeItem | Thenable<TreeItem> {
        return element;
    }

    // Returns the UI representation (AppItem) of the element that gets displayed in the view
    public getChildren(element?: AppRegItem | undefined): ProviderResult<AppRegItem[]> {
        // Top level so return all applications
        if (element === undefined) {
            return this._treeData;
        }

        // Return application owner children
        if (element.contextValue === "OWNERS") {
            return this.getApplicationOwners(element);
        }

        // Nothing more specific so return static defined children
        return element.children;
    }

    // Returns the application registration that is the parent of the given element
    public getParentApplication(objectId: string): Application {
        const app: AppRegItem = this._treeData.filter(item => item.objectId === objectId)[0];
        return app.manifest!;
    }

    // Returns all applications depending on the user setting
    private async getApplications(filter?: string): Promise<Application[]> {

        const showAllApplications = workspace.getConfiguration("applicationregistrations").get("showAllApplications") as boolean;
        
        // If not show all then get only owned applications
        if (showAllApplications === false) {
            const totalApps = await this._graphClient.getApplicationsOwnedCount();
            const response = await this._graphClient.getApplicationsOwned(filter);
            this._statusBarAppCount = window.setStatusBarMessage(`$(${filter === undefined ? "filter" : "filter-filled"}) ${response.value.length} of ${totalApps["@odata.count"]} Applications`);
            return response.value;
        } else {
            // Otherwise get all applications
            const totalApps = await this._graphClient.getApplicationsAllCount();
            const response = await this._graphClient.getApplicationsAll(filter);
            this._statusBarAppCount = window.setStatusBarMessage(`$(${filter === undefined ? "filter" : "filter-filled"}) ${response.value.length} of ${totalApps["@odata.count"]} Applications`);
            return response.value;
        }
    }

    // Returns the application owners for the given application
    private async getApplicationOwners(element: AppRegItem): Promise<AppRegItem[]> {
        return await this._graphClient.getApplicationOwners(element.objectId!)
            .then((response) => {
                const owners: User[] = response.value;
                let appOwners: AppRegItem[] = [];
                owners.forEach(owner => {
                    appOwners.push(new AppRegItem({
                        label: owner.displayName!,
                        context: "OWNER",
                        icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                        objectId: element.objectId,
                        userId: owner.id!
                    }));
                });
                return appOwners;
            });
    }
}
