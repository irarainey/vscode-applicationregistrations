import { TreeItem, TreeItemCollapsibleState } from 'vscode';
import { AppRegItemParams } from "../interfaces/appRegItemParams";

// This is the data structure for the application registration tree view item
export class AppRegItem extends TreeItem {

    // Private properties
    private _children?: AppRegItem[];
    private _objectId?: string = "";
    private _appId?: string = "";
    private _userId?: string = "";
    private _value?: string = "";
    private _order?: number = 0;
    private _state?: boolean = false;
    private _resourceAppId?: string = "";
    private _resourceScopeId?: string = "";

    // Constructor for the application registration tree view item
    constructor(params: AppRegItemParams) {

        // Call the base constructor
        super(params.label, params.children === undefined ? TreeItemCollapsibleState.None : TreeItemCollapsibleState.Collapsed);

        // Set other base properties
        this.contextValue = params.context;
        this.command = params.command;
        this.iconPath = params.icon;
        this.tooltip = params.tooltip;

        // Set the remaining properties
        this._children = params.children;
        this._value = params.value;
        this._objectId = params.objectId;
        this._appId = params.appId;
        this._userId = params.userId;
        this._order = params.order;
        this._state = params.state;
        this._resourceAppId = params.resourceAppId;
        this._resourceScopeId = params.resourceScopeId;
    }

    // Public property getter for the children
    public get children() {
        return this._children;
    }

    // Public property getter for the object id
    public get objectId() {
        return this._objectId;
    }

    // Public property getter for the application id
    public get appId() {
        return this._appId;
    }

    // Public property getter for the user id
    public get userId() {
        return this._userId;
    }

    // Public property getter for the value
    public get value() {
        return this._value;
    }

    // Public property getter for the order
    public get order() {
        return this._order;
    }

    // Public property getter for the state
    public get state() {
        return this._state;
    }

    // Public property getter for the resource api application id
    public get resourceAppId() {
        return this._resourceAppId;
    }

    // Public property getter for the resource api id
    public get resourceScopeId() {
        return this._resourceScopeId;
    }
}