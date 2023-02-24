import { ThemeIcon, TreeItem, TreeItemCollapsibleState } from "vscode";
import { AppRegItemParams } from "../types/app-reg-item-params";

// This is the data structure for the application registration tree view item
export class AppRegItem extends TreeItem {
	objectId?: string = "";
	baseIcon?: string | ThemeIcon = "";
	children?: AppRegItem[];
	appId?: string = "";
	userId?: string = "";
	keyId?: string = "";
	value?: string = "";
	order?: number = 0;
	state?: boolean = false;
	resourceAppId?: string = "";
	resourceScopeId?: string = "";

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
