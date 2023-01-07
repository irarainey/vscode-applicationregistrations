import { Disposable, Uri, ThemeIcon } from "vscode";
import { AppRegItem } from "../models/app-reg-item";

// This is the interface for the return result of service methods
export interface ActivityResult {
    success: boolean;
    error?: Error;
    statusBarHandle?: Disposable;
    treeViewItem?: AppRegItem;
    previousIcon?: string | Uri | {
        light: string | Uri;
        dark: string | Uri;
    } | ThemeIcon | undefined;
}