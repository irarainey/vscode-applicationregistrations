import { ThemeIcon, TreeItem, TreeItemCollapsibleState, TreeItemLabel } from "vscode";
import { AppRegItemParams } from "../types/app-reg-item-params";

// This is the data structure for the application registration tree view item
export class AppRegItem extends TreeItem {
	// Private properties
	public objectId?: string = "";
	public baseIcon?: string | ThemeIcon = "";
	public children?: AppRegItem[];
	public appId?: string = "";
	public userId?: string = "";
	public keyId?: string = "";
	public value?: string = "";
	public order?: number = 0;
	public state?: boolean = false;
	public resourceAppId?: string = "";
	public resourceScopeId?: string = "";
	public label?: string | TreeItemLabel | undefined;

	// Constructor for the application registration tree view item
	constructor(params: AppRegItemParams) {
		// Call the base constructor
		super(params.label, params.children === undefined ? TreeItemCollapsibleState.None : TreeItemCollapsibleState.Collapsed);

		// Set other base properties
		this.contextValue = params.context;
		this.command = params.command;
		this.iconPath = params.iconPath;
		this.tooltip = params.tooltip;
		this.label = params.label;

		// Set the remaining properties
		this.objectId = params.objectId;
		this.baseIcon = params.baseIcon;
		this.children = params.children;
		this.value = params.value;
		this.appId = params.appId;
		this.userId = params.userId;
		this.keyId = params.keyId;
		this.order = params.order;
		this.state = params.state;
		this.resourceAppId = params.resourceAppId;
		this.resourceScopeId = params.resourceScopeId;
	}
}
