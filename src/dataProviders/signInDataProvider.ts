import * as vscode from 'vscode';
import { signInCommandText } from '../constants';
import { TreeItem } from 'vscode';

// This is the data provider for the tree view when the user is not signed in to Azure CLI.
export class SignInDataProvider implements vscode.TreeDataProvider<TreeItem> {

    // This is the event that is fired when the tree view is refreshed.
    // Returns the children for the given element or root (if no element is passed).
    public getChildren(): Thenable<TreeItem[]> {
        return Promise.resolve([{label: "SIGNIN"}]);
    };

    // Returns the UI representation (TreeItem) of the element that gets displayed in the view
    public getTreeItem(_element: TreeItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return {
            label: signInCommandText,
            contextValue: "SIGNIN",
            iconPath: new vscode.ThemeIcon('sign-in'),
            command: {
                command: "appRegistrations.signInToAzure",
                title: signInCommandText,
            }
        };
    }

}
