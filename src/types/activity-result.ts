import { Disposable, Uri, ThemeIcon } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";

// This is the type for the return result of service methods
export type ActivityResult = {
    success: boolean;
    error?: Error;
    statusBarHandle?: Disposable;
    treeViewItem?: AppRegItem;
    treeDataProvider?: AppRegTreeDataProvider;
    previousIcon?: string | Uri | {
        light: string | Uri;
        dark: string | Uri;
    } | ThemeIcon | undefined;
};