import * as vscode from 'vscode';
import { view, portalUri } from './constants';
import { GraphClient } from './graphClient';
import { LoadingDataProvider } from './dataProviders/loadingDataProvider';
import { AppRegDataProvider, AppItem } from './dataProviders/appRegDataProvider';

export class ApplicationRegistrations {

    private filterCommand?: string = undefined;
    private filterText: string = '';
    private graphClient: GraphClient = new GraphClient();
    private subscriptions: vscode.Disposable[] = [];
    private authenticated: boolean = false;

    constructor({ subscriptions }: vscode.ExtensionContext) {
        vscode.commands.registerCommand(`${view}.addApp`, () => this.addApp());
        vscode.commands.registerCommand(`${view}.deleteApp`, node => this.deleteApp(node));
        vscode.commands.registerCommand(`${view}.renameApp`, node => this.renameApp(node));
        vscode.commands.registerCommand(`${view}.refreshApps`, () => this.refreshApps());
        vscode.commands.registerCommand(`${view}.filterApps`, () => this.filterApps());
        vscode.commands.registerCommand(`${view}.viewAppManifest`, node => this.viewAppManifest(node));
        vscode.commands.registerCommand(`${view}.copyAppId`, node => this.copyAppId(node));
        vscode.commands.registerCommand(`${view}.openAppInPortal`, node => this.openAppInPortal(node));
        vscode.commands.registerCommand(`${view}.copyValue`, node => this.copyValue(node));
        this.subscriptions = subscriptions;

        vscode.window.registerTreeDataProvider(view, new LoadingDataProvider());

        this.graphClient.authenticated = () => {
            this.authenticated = true;
            this.refreshApps();
        };
    }

    private refreshApps(): void {
        
        if (!this.authenticated) {
            return;
        }

        vscode.window.registerTreeDataProvider(view, new LoadingDataProvider());
        this.graphClient.getApplicationsAll(this.filterCommand)
            .then((apps) => {
                vscode.window.registerTreeDataProvider(view, new AppRegDataProvider(apps));
            }).catch((error) => {
                console.error(error);
            });
    };

    private async filterApps(): Promise<void> {

        if (!this.authenticated) {
            return;
        }

        const filterText = await vscode.window.showInputBox({
            placeHolder: "Name starts with...",
            prompt: "Filter applications by display name",
            value: this.filterText
        });

        if (filterText === '' || filterText === undefined) {
            this.filterCommand = undefined;
            this.filterText = '';
            this.refreshApps();
        } else {
            this.filterText = filterText;
            this.filterCommand = `startsWith(displayName, \'${filterText}\')`;
            this.refreshApps();
        }
    };

    private async addApp(): Promise<void> {

        if (!this.authenticated) {
            return;
        }

        const newName = await vscode.window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration with display name",
        });

        if (newName !== undefined) {
            this.graphClient.createApplication({ displayName: newName })
                .then((response) => {
                    console.log(response);
                    this.refreshApps();
                }).catch((error) => {
                    console.error(error);
                });
        }
    };

    private async renameApp(node: AppItem): Promise<void> {
        let app = node.manifest;

        const newName = await vscode.window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app!.displayName!
        });

        if (newName !== undefined) {
            this.graphClient.updateApplication(node.objectId!, { displayName: newName })
                .then((response) => {
                    this.refreshApps();
                }).catch((error) => {
                    console.error(error);
                });
        }

    };

    private deleteApp(node: AppItem): void {
        vscode.window
            .showInformationMessage(`Do you want to delete the application ${node.label}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    this.graphClient.deleteApplication(node.objectId!)
                        .then((response) => {
                            console.log(response);
                            this.refreshApps();
                        }).catch((error) => {
                            console.error(error);
                        });
                }
            });

    };

    private copyAppId(node: AppItem): void {
        vscode.env.clipboard.writeText(node.appId!);
        vscode.window.showInformationMessage(`Application Id: ${node.appId}`);
    };

    private copyValue(node: AppItem): void {
        vscode.env.clipboard.writeText(node.value!);
        vscode.window.showInformationMessage(`Value: ${node.label}`);
    };

    private openAppInPortal(node: AppItem): void {
        vscode.env.openExternal(vscode.Uri.parse(`${portalUri}${node.appId}`));
    }

    private async viewAppManifest(node: AppItem): Promise<void> {
        const myProvider = new class implements vscode.TextDocumentContentProvider {
            onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(uri: vscode.Uri): string {
                return JSON.stringify(node.manifest, null, 4);
            }
        };
        this.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider('manifest', myProvider));

        const uri = vscode.Uri.parse('manifest:' + node.label + ".json");
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.window.showTextDocument(doc, { preview: false });
    };
}