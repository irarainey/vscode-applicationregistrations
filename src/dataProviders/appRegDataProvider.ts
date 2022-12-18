import * as vscode from 'vscode';
import * as path from 'path';
import { Application } from "@microsoft/microsoft-graph-types";

export class AppRegDataProvider implements vscode.TreeDataProvider<AppItem> {
    
    private treeData: AppItem[] = [];

    constructor(apps: Application[]) {
        apps.forEach(app => {
            this.treeData.push(new AppItem({
                label: app.displayName ? app.displayName : "Application",
                context: "APPLICATION",
                objectId: app.id!,
                appId: app.appId!,
                manifest: app,
                children: [
                    new AppItem({
                        label: "Application Id",
                        context: "PROPERTY",
                        children:
                            [
                                new AppItem({
                                    label: app.appId!,
                                    context: "VALUE"
                                })
                            ]
                    }
                    )
                ]
            }
            ));
        });
    }

    public getTreeItem(element: AppItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return element;
    }

    public getChildren(element?: AppItem | undefined): vscode.ProviderResult<AppItem[]> {
        if (element === undefined) {
            return this.treeData;
        }
        return element.children;
    }
}

export class AppItem extends vscode.TreeItem {
    
    public children: AppItem[] | undefined;
    public context: string = "";
    public objectId?: string = "";
    public appId?: string = "";
    public manifest?: Application = {};
    public value?: string = "";

    constructor(params: AppParams) {
        super(
            params.label,
            params.children === undefined ? vscode.TreeItemCollapsibleState.None :
                vscode.TreeItemCollapsibleState.Collapsed);

        this.children = params.children;
        this.contextValue = params.context;
        this.value = params.value;
        this.objectId = params.objectId;
        this.appId = params.appId;
        this.manifest = params.manifest;

        switch (params.context) {
            case "APPLICATION":
                this.iconPath = path.join(__filename, '..', '..', '..', 'resources', "icons", "app.svg");
                break;
            case "PROPERTY":
                this.iconPath = {
                    light: path.join(__filename, '..', '..', '..', 'resources', "icons", "light", "property.svg"),
                    dark: path.join(__filename, '..', '..', '..', 'resources', "icons", "dark", "property.svg")
                };
                break;
            case "VALUE":
                this.iconPath = {
                    light: path.join(__filename, '..', '..', '..', 'resources', "icons", "light", "string.svg"),
                    dark: path.join(__filename, '..', '..', '..', 'resources', "icons", "dark", "string.svg")
                };
                break;
        }
    }
}

interface AppParams {
    label: string;
    context: string;
    value?: string;
    objectId?: string;
    appId?: string;
    manifest?: Application;
    children?: AppItem[];
}