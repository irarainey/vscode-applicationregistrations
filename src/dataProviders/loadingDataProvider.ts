import * as vscode from 'vscode';
import { TreeItem } from 'vscode';

// This is the data provider for the tree view when the data is being loaded.
export class LoadingDataProvider implements vscode.TreeDataProvider<TreeItem> {

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getChildren(): Thenable<TreeItem[]> {
        return Promise.resolve([{label: "LOADING"}]);
    };

    // Returns the UI representation (TreeItem) of the element that gets displayed in the view
    public getTreeItem(_element: TreeItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return {
            label: "Loading...",
            contextValue: "LOADING",
            iconPath: new vscode.ThemeIcon('loading~spin')
        };
    }

}