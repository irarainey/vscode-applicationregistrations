import * as vscode from 'vscode';
import { TreeItem } from 'vscode';

export class SignInDataProvider implements vscode.TreeDataProvider<TreeItem> {
    
    public getChildren(): Thenable<any[]> {
        return Promise.resolve(["SIGNIN"]);
    };

    public getTreeItem(_element: TreeItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return {
            label: <any>{ label: "Sign in to Azure CLI..." },
            contextValue: "SIGNIN",
            iconPath: new vscode.ThemeIcon('sign-in')
        };
    }
}