import * as path from 'path';
import { signInCommandText } from '../constants';
import { ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, Disposable } from 'vscode';
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphClient } from '../clients/graph';
import { AppRegItem } from '../models/appRegItem';

// This is the application registration data provider for the tree view.
export class AppRegDataProvider implements TreeDataProvider<AppRegItem> {

    // A private instance of the GraphClient class.
    private graphClient?: GraphClient;

    // Private instance of the tree data
    private treeData: AppRegItem[] = [];

    private statusBarMessage: Disposable | undefined;

    // This is the event that is fired when the tree view is refreshed.
    private _onDidChangeTreeData: EventEmitter<AppRegItem | undefined | null | void> = new EventEmitter<AppRegItem | undefined | null | void>();

    //Defines the event that is fired when the tree view is refreshed.
    readonly onDidChangeTreeData: Event<AppRegItem | undefined | null | void> = this._onDidChangeTreeData.event;

    // Initialises the tree view data based on the type of data to be displayed.
    public initialise(type: string, statusBarMessage: Disposable | undefined, graphClient?: GraphClient, apps?: Application[]) {

        // Clear any existing status bar message
        if(this.statusBarMessage !== undefined) {
            this.statusBarMessage.dispose();
        }
        
        this.statusBarMessage = statusBarMessage;

        if (graphClient !== undefined) {
            this.graphClient = graphClient;
        }

        // Clear the tree data
        this.treeData = [];

        // Add the appropriate tree view item based on the type of data to be displayed.
        switch (type) {
            case "LOADING":
                this.treeData.push(new AppRegItem({
                    label: "Loading...",
                    context: "LOADING",
                    icon: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
                }));
                break;
            case "SIGN-IN":
                this.treeData.push(new AppRegItem({
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
                this.buildAppRegTree(apps!);
                break;
        }

        // Fire the event to refresh the tree view
        this._onDidChangeTreeData.fire();
    }

    // Trigger the event to refresh the tree view
    public triggerOnDidChangeTreeData() {
        this._onDidChangeTreeData.fire();
    }

    private buildAppRegTree(apps: Application[]) {
        // Iterate through the applications and create the tree data
        apps!.forEach(app => {
            // Create the tree view item for the application and it's children
            this.treeData.push(new AppRegItem({
                label: app.displayName ? app.displayName : "Application",
                context: "APPLICATION",
                icon: path.join(__filename, "..", "..", "..","resources", "icons", "app.svg"),
                objectId: app.id!,
                appId: app.appId!,
                manifest: app,
                children: [
                    // Application Id
                    new AppRegItem({
                        label: "Application Id",
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
                                children: app.web?.redirectUris?.length === 0 ? undefined : app.web?.redirectUris?.map(uri => {
                                    return new AppRegItem({
                                        label: uri,
                                        context: "WEB-REDIRECT-URI",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                    });
                                })
                            }),
                            new AppRegItem({
                                label: "SPA",
                                context: "SPA-REDIRECT",
                                icon: new ThemeIcon("browser"),
                                objectId: app.id!,
                                children: app.spa?.redirectUris?.length === 0 ? undefined : app.spa?.redirectUris?.map(uri => {
                                    return new AppRegItem({
                                        label: uri,
                                        context: "SPA-REDIRECT-URI",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                    });
                                })
                            }),
                            new AppRegItem({
                                label: "Mobile and Desktop",
                                context: "NATIVE-REDIRECT",
                                icon: new ThemeIcon("editor-layout"),
                                objectId: app.id!,
                                children: app.publicClient?.redirectUris?.length === 0 ? undefined : app.publicClient?.redirectUris?.map(uri => {
                                    return new AppRegItem({
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
                    new AppRegItem({
                        label: "Credentials",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                        children: [
                            new AppRegItem({
                                label: "Client Secrets",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
                                children: app.passwordCredentials!.length === 0 ? undefined : app.passwordCredentials!.map(password => {
                                    return new AppRegItem({
                                        label: password.displayName!,
                                        context: "PASSWORD",
                                        icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        value: password.keyId!,
                                        children: [
                                            new AppRegItem({
                                                label: "Created",
                                                context: "PASSWORD-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
                                                        label: password.startDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppRegItem({
                                                label: "Expires",
                                                context: "PASSWORD-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
                                                        label: password.endDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            })
                                        ]
                                    });
                                })
                            }),
                            new AppRegItem({
                                label: "Certificates",
                                context: "PROPERTY-ARRAY",
                                icon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
                                children: app.keyCredentials!.length === 0 ? undefined : app.keyCredentials!.map(password => {
                                    return new AppRegItem({
                                        label: password.displayName!,
                                        context: "CERTIFICATE",
                                        icon: new ThemeIcon("symbol-key", new ThemeColor("editor.foreground")),
                                        objectId: app.id!,
                                        value: password.keyId!,
                                        children: [
                                            new AppRegItem({
                                                label: "Type",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
                                                        label: password.type!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppRegItem({
                                                label: "Key Identifier",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
                                                        label: password.customKeyIdentifier!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppRegItem({
                                                label: "Created",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
                                                        label: password.startDateTime!,
                                                        context: "PASSWORD-VALUE",
                                                    }),
                                                ]
                                            }),
                                            new AppRegItem({
                                                label: "Expires",
                                                context: "CERTIFICATE-VALUE",
                                                icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                                children: [
                                                    new AppRegItem({
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
                    new AppRegItem({
                        label: "API Permissions",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
                        children: app.requiredResourceAccess?.length === 0 ? undefined : app.requiredResourceAccess!.map(api => {
                            return new AppRegItem({
                                label: api.resourceAppId!,
                                context: "SCOPE",
                                icon: new ThemeIcon("symbol-variable", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: api.resourceAppId!,
                                children: api.resourceAccess!.map(scope => {
                                    return new AppRegItem({
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
                    new AppRegItem({
                        label: "Exposed API Permissions",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
                        children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : app.api!.oauth2PermissionScopes!.map(scope => {
                            return new AppRegItem({
                                label: scope.adminConsentDisplayName!,
                                context: "SCOPE",
                                icon: new ThemeIcon("symbol-variable", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: scope.id!,
                                children: [
                                    new AppRegItem({
                                        label: scope.value!,
                                        context: "SCOPE-VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppRegItem({
                                        label: scope.adminConsentDescription!,
                                        context: "SCOPE-DESCRIPTION",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppRegItem({
                                        label: scope.type! === "User" ? "Admins and Users" : "Admins Only",
                                        context: "SCOPE-TYPE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppRegItem({
                                        label: scope.isEnabled! ? "Enabled" : "Disabled",
                                        context: "SCOPE-ENABLED",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    })
                                ]
                            });
                        })
                    }),
                    // App Roles
                    new AppRegItem({
                        label: "App Roles",
                        context: "PROPERTY-ARRAY",
                        icon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
                        children: app.appRoles!.length === 0 ? undefined : app.appRoles!.map(role => {
                            return new AppRegItem({
                                label: role.displayName!,
                                context: "ROLE",
                                icon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
                                objectId: app.id!,
                                value: role.id!,
                                children: [
                                    new AppRegItem({
                                        label: role.value!,
                                        context: "ROLE-VALUE",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppRegItem({
                                        label: role.description!,
                                        context: "ROLE-DESCRIPTION",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    }),
                                    new AppRegItem({
                                        label: role.isEnabled! ? "Enabled" : "Disabled",
                                        context: "ROLE-ENABLED",
                                        icon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
                                    })
                                ]
                            });
                        })
                    }),
                    // Owners
                    new AppRegItem({
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

        // Clear any status bar message
        if(this,this.statusBarMessage !== undefined) {
            this.statusBarMessage.dispose();
        }
    }

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getTreeItem(element: AppRegItem): TreeItem | Thenable<TreeItem> {
        return element;
    }

    // Returns the UI representation (AppItem) of the element that gets displayed in the view
    public getChildren(element?: AppRegItem | undefined): ProviderResult<AppRegItem[]> {
        if (element === undefined) {
            return this.treeData;
        }
        return element.children;
    }

    // Returns the application registration that is the parent of the given element
    public async getParentApplication(objectId: string): Promise<Application> {
        const app: AppRegItem = this.treeData.filter(item => item.objectId === objectId)[0];
        return app.manifest!;
    }

    // Returns the owners of the application registration as an array of AppItem
    private getApplicationOwners(objectId: string): AppRegItem[] {

        let appOwners: AppRegItem[] = [];

        this.graphClient?.getApplicationOwners(objectId)
            .then(owners => {
                owners.forEach(owner => {
                    appOwners.push(new AppRegItem({
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
