import * as vscode from 'vscode';
import { execShellCmd } from './utils';
import { view, portalUri } from './constants';
import { GraphClient } from './graphClient';
import { SignInDataProvider } from './dataProviders/signInDataProvider';
import { LoadingDataProvider } from './dataProviders/loadingDataProvider';
import { AppRegDataProvider, AppItem } from './dataProviders/appRegDataProvider';

export class ApplicationRegistrations {

    private graphClient: GraphClient;
    private filterCommand?: string = undefined;
    private filterText: string = '';
    private subscriptions: vscode.Disposable[] = [];
    private authenticated: boolean = false;
    public isUserAuthenticated: (state: boolean | undefined) => void;

    constructor(graphClient: GraphClient, { subscriptions }: vscode.ExtensionContext) {
        this.graphClient = graphClient;
        this.subscriptions = subscriptions;
        this.isUserAuthenticated = () => { };
        vscode.window.registerTreeDataProvider(view, new LoadingDataProvider());
        this.determineAuthenticationState();
    }

    private determineAuthenticationState(): void {
        if (this.authenticated === false) {
            this.graphClient.initialise();
            this.graphClient.authenticationStateChange = (state: boolean | undefined) => {
                if (state === true) {
                    this.authenticated = true;
                    this.populateTreeView();
                } else if (state === false) {
                    vscode.window.registerTreeDataProvider(view, new SignInDataProvider());
                    this.isUserAuthenticated = (state: boolean | undefined) => {
                        if (state === true) {
                            this.determineAuthenticationState();
                        } else if (state === false) {
                            vscode.window.showErrorMessage("Please sign in to Azure CLI.");
                        }
                    };
                }
            };
        } else {
            this.populateTreeView();
        }
    }

    public populateTreeView(): void {
        vscode.window.registerTreeDataProvider(view, new LoadingDataProvider());
        this.graphClient.getApplicationsAll(this.filterCommand)
            .then((apps) => {
                vscode.window.registerTreeDataProvider(view, new AppRegDataProvider(apps));
            }).catch(() => {
                this.authenticated = false;
                this.determineAuthenticationState();
            });
    };

    public async filterApps(): Promise<void> {

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
            this.populateTreeView();
        } else {
            this.filterText = filterText;
            this.filterCommand = `startsWith(displayName, \'${filterText}\')`;
            this.populateTreeView();
        }
    };

    public async addApp(): Promise<void> {

        if (!this.authenticated) {
            return;
        }

        const newName = await vscode.window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration with display name",
        });

        if (newName !== undefined) {
            this.graphClient.createApplication({ displayName: newName })
                .then(() => {
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }
    };

    public async renameApp(app: AppItem): Promise<void> {
        const newName = await vscode.window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app.manifest!.displayName!
        });

        if (newName !== undefined) {
            this.graphClient.updateApplication(app.objectId!, { displayName: newName })
                .then(() => {
                    this.populateTreeView();
                }).catch((error) => {
                    console.error(error);
                });
        }

    };

    public deleteApp(app: AppItem): void {
        vscode.window
            .showInformationMessage(`Do you want to delete the application ${app.label}?`, "Yes", "No")
            .then(answer => {
                if (answer === "Yes") {
                    this.graphClient.deleteApplication(app.objectId!)
                        .then((response) => {
                            console.log(response);
                            this.populateTreeView();
                        }).catch((error) => {
                            console.error(error);
                        });
                }
            });

    };

    public copyAppId(app: AppItem): void {
        vscode.env.clipboard.writeText(app.appId!);
        vscode.window.showInformationMessage(`Application Id: ${app.appId}`);
    };

    public openAppInPortal(app: AppItem): void {
        vscode.env.openExternal(vscode.Uri.parse(`${portalUri}${app.appId}`));
    }

    public async viewAppManifest(app: AppItem): Promise<void> {
        const myProvider = new class implements vscode.TextDocumentContentProvider {
            onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(uri: vscode.Uri): string {
                return JSON.stringify(app.manifest, null, 4);
            }
        };
        this.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider('manifest', myProvider));

        const uri = vscode.Uri.parse('manifest:' + app.label + ".json");
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.window.showTextDocument(doc, { preview: false });
    };

    public copyValue(item: AppItem): void {
        vscode.env.clipboard.writeText(item.value!);
        vscode.window.showInformationMessage(`Value: ${item.label}`);
    };

    public async invokeSignIn(): Promise<void> {

        const tenant = await vscode.window.showInputBox({
            placeHolder: "Tenant name or Id...",
            prompt: "Enter the tenant name or Id, or leave blank for the default tenant",
        });

        if(tenant === undefined) {
            return;
        } 

        let command = "az login";
        if(tenant.length > 0) { 
            command += ` --tenant ${tenant}`;
        }

        execShellCmd(command)
            .then(() => {
                this.isUserAuthenticated(true);
            }).catch(() => {
                this.isUserAuthenticated(false);
            });
    }
}