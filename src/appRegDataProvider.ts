import * as vscode from 'vscode';
import * as path from 'path';
import { signInCommandText } from './constants';
import { ThemeIcon, ThemeColor } from 'vscode';
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphClient } from './graphClient';

// This is the data provider for the tree view.
export class AppRegDataProvider implements vscode.TreeDataProvider<AppItem> {

    // A private instance of the GraphClient class.
    private graphClient?: GraphClient;

    // Private instance of the tree data
    private treeData: AppItem[] = [];

    // This is the event that is fired when the tree view is refreshed.
    private _onDidChangeTreeData: vscode.EventEmitter<AppItem | undefined | null | void> = new vscode.EventEmitter<AppItem | undefined | null | void>();

    //Defines the event that is fired when the tree view is refreshed.
    readonly onDidChangeTreeData: vscode.Event<AppItem | undefined | null | void> = this._onDidChangeTreeData.event;

    // Initialises the tree view data based on the type of data to be displayed.
    public initialise(type: string, graphClient?: GraphClient, apps?: Application[]) {

        if (graphClient !== undefined) {
            this.graphClient = graphClient;
        }

        // Clear the tree data
        this.treeData = [];

        // Add the appropriate tree view item based on the type of data to be displayed.
        switch (type) {
            case "LOADING":
                this.treeData.push(new AppItem({
                    label: "Loading...",
                    context: "LOADING",
                    icon: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
                }));
                break;
            case "SIGN-IN":
                this.treeData.push(new AppItem({
                    label: signInCommandText,
                    context: "SIGN-IN",
                    icon: new ThemeIcon("sign-in", new ThemeColor("editor.foreground")),
                    command: {
                        command: "appRegistrations.signInToAzure",
                        title: signInCommandText,
                    }
                }));
                break;
            case "APPLICATIONS":
                this.buildTree(apps!);
                break;
        }

        // Fire the event to refresh the tree view
        this._onDidChangeTreeData.fire();
    }

    // Trigger the event to refresh the tree view
    public triggerOnDidChangeTreeData() {
        this._onDidChangeTreeData.fire();
    }

    private buildTree(apps: Application[]) {
        // Iterate through the applications and create the tree data
        apps!.forEach(app => {

            if(app.displayName!.startsWith("aad-extensions-app")){
                return;
            }

            // Create the tree view item for the application and it's children
            this.treeData.push(new AppItem({
                label: app.displayName ? app.displayName : "Application",
                context: "APPLICATION",
                icon: path.join(__filename, "..", "..", "resources", "icons", "app.svg"),
                objectId: app.id!,
                appId: app.appId!,
                manifest: app,
                children: [
                    // Application Id
                    new AppItem({
                        label: "Application Id",
                        context: "APPID-PARENT",
                        icon: new ThemeIcon("preview"),
                        children: [
                            new AppItem({
                                label: app.appId!,
                                value: app.appId!,
                                context: "COPY",
                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                            })
                        ]
                    }),
                    // Sign In Audience
                    new AppItem({
                        label: "Sign In Audience",
                        context: "AUDIENCE-PARENT",
                        icon: new ThemeIcon("account"),
                        objectId: app.id!,
                        children: [
                            new AppItem({
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
                    new AppItem({
                        label: "Redirect URIs",
                        context: "REDIRECT-PARENT",
                        icon: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
                        objectId: app.id!,
                        children: [
                            new AppItem({
                                label: "Web",
                                context: "WEB-REDIRECT",
                                icon: new ThemeIcon("globe"),
                                objectId: app.id!,
                                children: app.web?.redirectUris?.length === 0 ? undefined : app.web?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "WEB-REDIRECT-URI",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                    });
                                })
                            }),
                            new AppItem({
                                label: "SPA",
                                context: "SPA-REDIRECT",
                                icon: new ThemeIcon("browser"),
                                objectId: app.id!,
                                children: app.spa?.redirectUris?.length === 0 ? undefined : app.spa?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "SPA-REDIRECT-URI",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                    });
                                })
                            }),
                            new AppItem({
                                label: "Mobile and Desktop",
                                context: "NATIVE-REDIRECT",
                                icon: new ThemeIcon("editor-layout"),
                                objectId: app.id!,
                                children: app.publicClient?.redirectUris?.length === 0 ? undefined : app.publicClient?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "NATIVE-REDIRECT-URI",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                    });
                                })
                            })
                        ]
                    }),
                    // Credentials
                    new AppItem({
                        label: "Credentials",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                        children: [
                            new AppItem({
                                label: "Client Secrets",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                children: app.passwordCredentials!.length === 0 ? undefined : app.passwordCredentials!.map(password => {
                                    return new AppItem({
                                        label: password.displayName!,
                                        context: "PASSWORD",
                                        icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        value: password.keyId!,
                                        children: [
                                            new AppItem({
                                                label: "Created",
                                                context: "PASSWORD-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.startDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppItem({
                                                label: "Expires",
                                                context: "PASSWORD-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.endDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            })
                                        ]
                                    });
                                })
                            }),
                            new AppItem({
                                label: "Certificates",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                children: app.keyCredentials!.length === 0 ? undefined : app.keyCredentials!.map(password => {
                                    return new AppItem({
                                        label: password.displayName!,
                                        context: "CERTIFICATE",
                                        icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        value: password.keyId!,
                                        children: [
                                            new AppItem({
                                                label: "Type",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.type!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppItem({
                                                label: "Key Identifier",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.customKeyIdentifier!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppItem({
                                                label: "Created",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.startDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppItem({
                                                label: "Expires",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppItem({
                                                        label: password.endDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            })
                                        ]
                                    });
                                })

                            })
                        ]
                    }),
                    // API Permissions
                    new AppItem({
                        label: "API Permissions",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                        children: app.requiredResourceAccess?.length === 0 ? undefined : app.requiredResourceAccess!.map(api => {
                            return new AppItem({
                                label: api.resourceAppId!,
                                context: "SCOPE",
                                icon: new ThemeIcon("symbol-variable", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: api.resourceAppId!,
                                children: api.resourceAccess!.map(scope => {
                                    return new AppItem({
                                        label: scope.type!,
                                        context: "SCOPE-VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        value: scope.id!,
                                    });
                                })
                            });
                        })
                    }),
                    // Exposed API Permissions
                    new AppItem({
                        label: "Exposed API Permissions",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                        children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : app.api!.oauth2PermissionScopes!.map(scope => {
                            return new AppItem({
                                label: scope.adminConsentDisplayName!,
                                context: "SCOPE",
                                icon: new ThemeIcon("symbol-variable", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: scope.id!,
                                children: [
                                    new AppItem({
                                        label: scope.value!,
                                        context: "SCOPE-VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppItem({
                                        label: scope.adminConsentDescription!,
                                        context: "SCOPE-DESCRIPTION",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppItem({
                                        label: scope.type! === "User" ? "Admins and Users" : "Admins Only",
                                        context: "SCOPE-TYPE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppItem({
                                        label: scope.isEnabled! ? "Enabled" : "Disabled",
                                        context: "SCOPE-ENABLED",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    })
                                ]
                            });
                        })
                    }),
                    // App Roles
                    new AppItem({
                        label: "App Roles",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                        children: app.appRoles!.length === 0 ? undefined : app.appRoles!.map(role => {
                            return new AppItem({
                                label: role.displayName!,
                                context: "ROLE",
                                icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: role.id!,
                                children: [
                                    new AppItem({
                                        label: role.value!,
                                        context: "ROLE-VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppItem({
                                        label: role.description!,
                                        context: "ROLE-DESCRIPTION",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppItem({
                                        label: role.isEnabled! ? "Enabled" : "Disabled",
                                        context: "ROLE-ENABLED",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    })
                                ]
                            });
                        })
                    }),
                    // Owners
                    new AppItem({
                        label: "Owners",
                        context: "OWNERS",
                        icon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
                        objectId: app.id!,
                        children: this.getApplicationOwners(app.id!)
                    }),
                ]
            }
            ));
        });
    }

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getTreeItem(element: AppItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return element;
    }

    // Returns the UI representation (AppItem) of the element that gets displayed in the view
    public getChildren(element?: AppItem | undefined): vscode.ProviderResult<AppItem[]> {
        if (element === undefined) {
            return this.treeData;
        }
        return element.children;
    }

    // Returns the application registration that is the parent of the given element
    public getParentApplication(objectId: string): Application {
        const app: AppItem = this.treeData.filter(item => item.objectId === objectId)[0];
        return app.manifest!;
    }

    // Returns the owners of the application registration as an array of AppItem
    private getApplicationOwners(objectId: string): AppItem[] {

        let appOwners: AppItem[] = [];

        this.graphClient?.getApplicationOwners(objectId)
            .then(owners => {
                owners.forEach(owner => {
                    appOwners.push(new AppItem({
                        label: owner.displayName!,
                        context: "OWNER",
                        icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                        objectId: objectId,
                        userId: owner.id!
                    }));
                });
            });

        return appOwners;

    }

}

// This is the data structure for the application registration tree view item
export class AppItem extends vscode.TreeItem {

    // Public properties
    public children: AppItem[] | undefined;
    public context: string = "";
    public objectId?: string = "";
    public appId?: string = "";
    public userId?: string = "";
    public manifest?: Application = {};
    public value?: string = "";

    // Constructor
    constructor(params: AppParams) {

        // Call the base constructor
        super(params.label, params.children === undefined ? vscode.TreeItemCollapsibleState.None : vscode.TreeItemCollapsibleState.Collapsed);

        // Set the remaining properties
        this.children = params.children;
        this.contextValue = params.context;
        this.value = params.value;
        this.objectId = params.objectId;
        this.appId = params.appId;
        this.userId = params.userId;
        this.manifest = params.manifest;
        this.command = params.command;
        this.iconPath = params.icon;
    }

}

// This is the interface for the application registration tree view item parameters
interface AppParams {
    label: string;
    context: string;
    value?: string;
    objectId?: string;
    appId?: string;
    userId?: string;
    manifest?: Application;
    children?: AppItem[];
    command?: vscode.Command;
    icon?: string | ThemeIcon;
}