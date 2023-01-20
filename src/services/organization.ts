import { window, Uri, TextDocumentContentProvider, EventEmitter, workspace, Disposable } from "vscode";
import { CLI_TENANT_CMD } from "../constants";
import { ServiceBase } from "./service-base";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { execShellCmd } from "../utils/exec-shell-cmd";
import { GraphResult } from "../types/graph-result";
import { Organization } from "@microsoft/microsoft-graph-types";

export class OrganizationService extends ServiceBase {

    // The constructor for the OwnerService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Shows the tenant information.
    async showTenantInformation(): Promise<void> {

        // Check if the graph client is initialised.
        if (this.graphRepository.isClientInitialised === false) {
            await this.treeDataProvider.initialiseGraphClient();
            return;
        }

        // Set the status bar message.
        this.triggerTreeChange("Loading Tenant Information");

        // Execute the az cli command to get the tenant id
        execShellCmd(CLI_TENANT_CMD)
            .then(async (response) => {
                await this.showTenantWindow(response);
            })
            .catch(async (error) => {
                this.triggerOnError(error);
            });
    }

    // Shows the tenant information in a new read-only window.
    private async showTenantWindow(tenantId: string): Promise<void> {
        // Get the tenant information.
        const result: GraphResult<Organization> = await this.graphRepository.getTenantInformation(tenantId);
        if (result.success === true && result.value !== undefined) {
            const tenantInformation = {
                id: result.value.id,
                displayName: result.value.displayName,
                primaryDomain: result.value.verifiedDomains?.filter(x => x.isDefault === true)[0].name,
                verifiedDomains: result.value.verifiedDomains?.map(x => x.name)
            };

            const newDocument = new class implements TextDocumentContentProvider {
                onDidChangeEmitter = new EventEmitter<Uri>();
                onDidChange = this.onDidChangeEmitter.event;
                provideTextDocumentContent(): string {
                    return JSON.stringify(tenantInformation, null, 4);
                }
            };

            this.disposable.push(workspace.registerTextDocumentContentProvider("tenantInformation", newDocument));
            const uri = Uri.parse(`tenantInformation:Tenant - ${result.value.displayName}.json`);
            workspace.openTextDocument(uri)
                .then(async (doc) => {
                    await window.showTextDocument(doc, { preview: false });
                    if (this.statusBarHandle !== undefined) {
                        this.statusBarHandle!.dispose();
                    }
                });
        } else {
            this.triggerOnError(result.error);
            return;
        }
    }
}