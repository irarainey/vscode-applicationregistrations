import { TreeItem, TreeItemCollapsibleState } from 'vscode';
import { Application } from "@microsoft/microsoft-graph-types";
import { AppRegItemParams } from "../interfaces/appRegItemParams";

// This is the data structure for the application registration tree view item
export class AppRegItem extends TreeItem {

    // Public properties
    public children: AppRegItem[] | undefined;
    public context: string = "";
    public objectId?: string = "";
    public appId?: string = "";
    public userId?: string = "";
    public manifest?: Application = {};
    public value?: string = "";

    // Constructor
    constructor(params: AppRegItemParams) {

        // Call the base constructor
        super(params.label, params.children === undefined ? TreeItemCollapsibleState.None : TreeItemCollapsibleState.Collapsed);

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