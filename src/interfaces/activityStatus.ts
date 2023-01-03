import { Disposable, Uri, ThemeIcon } from 'vscode';
import { AppRegItem } from '../models/appRegItem';

// This is the interface for the return result of service methods
export interface ActivityStatus {
    success: boolean;
    error?: Error;
    statusBarHandle?: Disposable;
    treeViewItem?: AppRegItem;
    previousIcon?: string | Uri | {
        light: string | Uri;
        dark: string | Uri;
    } | ThemeIcon | undefined;
}