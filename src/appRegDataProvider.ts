import * as vscode from 'vscode';
import * as path from 'path';
import { signInCommandText } from './constants';
import { ThemeIcon, ThemeColor } from 'vscode';
import { Application } from "@microsoft/microsoft-graph-types";

// This is the data provider for the tree view.
export class AppRegDataProvider implements vscode.TreeDataProvider<AppItem> {

    // Private instance of the tree data
    private treeData: AppItem[] = [];

    // This is the event that is fired when the tree view is refreshed.
    private _onDidChangeTreeData: vscode.EventEmitter<AppItem | undefined | null | void> = new vscode.EventEmitter<AppItem | undefined | null | void>();

    //Defines the event that is fired when the tree view is refreshed.
    readonly onDidChangeTreeData: vscode.Event<AppItem | undefined | null | void> = this._onDidChangeTreeData.event;

    // Initialises the tree view data based on the type of data to be displayed.
    public initialise(type: string, apps?: Application[]) {

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
            case "SIGNIN":
                this.treeData.push(new AppItem({
                    label: signInCommandText,
                    context: "SIGNIN",
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

    private buildTree(apps: Application[]) {
        // Iterate through the applications and create the tree data
        apps!.forEach(app => {
            // Create the tree view item for the application and it's children
            this.treeData.push(new AppItem({
                label: app.displayName ? app.displayName : "Application",
                context: "APPLICATION",
                icon: path.join(__filename, "..", "..", "resources", "icons", "app.svg"),
                objectId: app.id!,
                appId: app.appId!,
                manifest: app,
                children: [
                    new AppItem({
                        label: "Application Id",
                        context: "PROPERTY",
                        icon: new ThemeIcon("preview"),
                        children: [
                            new AppItem({
                                label: app.appId!,
                                context: "VALUE",
                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                            })
                        ]
                    }),
                    new AppItem({
                        label: "Sign In Audience",
                        context: "PROPERTY",
                        icon: new ThemeIcon("account"),
                        children: [
                            new AppItem({
                                label: app.signInAudience! === "AzureADMyOrg"
                                    ? "Single Tenant"
                                    : app.signInAudience! === "AzureADMultipleOrgs"
                                        ? "Multi Tenant"
                                        : "Multi Tenant and Personal Accounts",
                                context: "VALUE",
                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                            })
                        ]
                    }),
                    new AppItem({
                        label: "Redirect URIs",
                        context: "PROPERTY",
                        icon: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
                        children: [
                            new AppItem({
                                label: "Web",
                                context: "PROPERTY",
                                icon: new ThemeIcon("globe"),
                                children: app.web?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                                    });
                                })
                            }),
                            new AppItem({
                                label: "SPA",
                                context: "PROPERTY",
                                icon: new ThemeIcon("browser"),
                                children: app.spa?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                                    });
                                })
                            }),
                            new AppItem({
                                label: "Mobile and Desktop",
                                context: "PROPERTY",
                                icon: new ThemeIcon("editor-layout"),
                                children: app.publicClient?.redirectUris?.map(uri => {
                                    return new AppItem({
                                        label: uri,
                                        context: "VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"))
                                    });
                                })
                            })
                        ]
                    }),
                    new AppItem({
                        label: "Credentials",
                        context: "PROPERTY",
                        icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                        children: [
                            new AppItem({
                                label: "Client Secrets",
                                context: "PROPERTY",
                                icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                children: []
                            }),
                            new AppItem({
                                label: "Certificates",
                                context: "PROPERTY",
                                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                children: []
                            })
                        ]
                    }),
                    new AppItem({
                        label: "API Permissions",
                        context: "PROPERTY",
                        icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                        children: []
                    }),
                    new AppItem({
                        label: "Exposed API Permissions",
                        context: "PROPERTY",
                        icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                        children: []
                    }),
                    new AppItem({
                        label: "App Roles",
                        context: "PROPERTY",
                        icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                        children: []
                    }),
                    new AppItem({
                        label: "Owners",
                        context: "PROPERTY",
                        icon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
                        children: []
                    })
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

}

// This is the data structure for the application registration tree view item
export class AppItem extends vscode.TreeItem {

    // Public properties
    public children: AppItem[] | undefined;
    public context: string = "";
    public objectId?: string = "";
    public appId?: string = "";
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
    manifest?: Application;
    children?: AppItem[];
    command?: vscode.Command;
    icon?: string | ThemeIcon;
}