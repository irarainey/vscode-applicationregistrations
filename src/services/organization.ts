import { window, Uri, TextDocumentContentProvider, EventEmitter, workspace, Disposable } from "vscode";
import { CLI_TENANT_CMD } from "../constants";
import { ServiceBase } from "./service-base";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { GraphClient, execShellCmd } from "../clients/graph-client";

export class OrganizationService extends ServiceBase {

    // The constructor for the OwnerService class.
    constructor(graphClient: GraphClient, treeDataProvider: AppRegTreeDataProvider) {
        super(graphClient, treeDataProvider);
    }

    // Shows the tenant information.
    async showTenantInformation(): Promise<void> {

        // Check if the graph client is initialised.
        if (this.graphClient.isGraphClientInitialised === false) {
            await this.treeDataProvider.initialiseGraphClient();
            return;
        }

        // Set the status bar message.
        const status = this.triggerTreeChange("Loading Tenant Information");

        // Execute the az cli command to get the tenant id
        execShellCmd(CLI_TENANT_CMD)
            .then(async (response) => {
                await this.showTenantWindow(response, status);
            })
            .catch(async (error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error });
            });
    }

    // Shows the tenant information in a new read-only window.
    private async showTenantWindow(tenantId: string, statusBarHandle: Disposable | undefined): Promise<void> {
        // Get the tenant information.
        this.graphClient.getTenantInformation(tenantId)
            .then(async (response) => {
                const tenantInformation = {
                    id: response.id,
                    displayName: response.displayName,
                    primaryDomain: response.verifiedDomains?.filter(x => x.isDefault === true)[0].name,
                    verifiedDomains: response.verifiedDomains?.map(x => x.name)
                };

                const newDocument = new class implements TextDocumentContentProvider {
                    onDidChangeEmitter = new EventEmitter<Uri>();
                    onDidChange = this.onDidChangeEmitter.event;
                    provideTextDocumentContent(): string {
                        return JSON.stringify(tenantInformation, null, 4);
                    }
                };

                this.disposable.push(workspace.registerTextDocumentContentProvider('tenantInformation', newDocument));
                const uri = Uri.parse('tenantInformation:' + response.displayName + ".json");
                workspace.openTextDocument(uri)
                    .then(async (doc) => {
                        await window.showTextDocument(doc, { preview: false });
                        statusBarHandle!.dispose();
                    });
            })
            .catch((error) => {
                statusBarHandle!.dispose();
                this.triggerOnError({ success: false, statusBarHandle: statusBarHandle, error: error });
            });
    }
}