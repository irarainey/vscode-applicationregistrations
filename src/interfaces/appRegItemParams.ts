import { ThemeIcon, Command } from 'vscode';
import { Application } from "@microsoft/microsoft-graph-types";
import { AppRegItem } from "../models/appRegItem";

// This is the interface for the application registration tree view item parameters
export interface AppRegItemParams {
    label: string;
    context: string;
    value?: string;
    objectId?: string;
    appId?: string;
    userId?: string;
    manifest?: Application;
    children?: AppRegItem[];
    command?: Command;
    icon?: string | ThemeIcon;
    tooltip?: string;
}