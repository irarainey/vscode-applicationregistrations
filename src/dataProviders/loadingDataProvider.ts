import * as vscode from 'vscode';
import { TreeItem } from 'vscode';

export class LoadingDataProvider implements vscode.TreeDataProvider<TreeItem> {
    
    public getChildren(): Thenable<any[]> {
        return Promise.resolve(["LOADING"]);
    };

    public getTreeItem(_element: TreeItem): vscode.TreeItem | Thenable<vscode.TreeItem> {
        return {
            label: <any>{ label: "Loading..." },
            contextValue: "LOADING",
            iconPath: new vscode.ThemeIcon('loading~spin')
        };
    }
}