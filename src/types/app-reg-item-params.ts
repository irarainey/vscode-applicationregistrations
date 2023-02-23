import { ThemeIcon, Command } from "vscode";
import { AppRegItem } from "../models/app-reg-item";

// This is the type for the application registration tree view item parameters
export type AppRegItemParams = {
	label: string;
	context: string;
	iconPath: string | ThemeIcon;
	objectId?: string;
	baseIcon?: string | ThemeIcon;
	value?: string;
	appId?: string;
	resourceAppId?: string;
	resourceScopeId?: string;
	userId?: string;
	keyId?: string;
	children?: AppRegItem[];
	command?: Command;
	tooltip?: string;
	order?: number;
	state?: boolean;
};
