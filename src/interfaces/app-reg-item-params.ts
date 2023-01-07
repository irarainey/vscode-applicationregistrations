import { ThemeIcon, Command } from "vscode";
import { AppRegItem } from "../models/app-reg-item";

// This is the interface for the application registration tree view item parameters
export interface AppRegItemParams {
    label: string;
    context: string;
    value?: string;
    objectId?: string;
    appId?: string;
    resourceAppId?: string;
    resourceScopeId?: string;
    userId?: string;
    children?: AppRegItem[];
    command?: Command;
    icon?: string | ThemeIcon;
    tooltip?: string;
    order?: number;
    state?: boolean;
}