import { window, Uri, TextDocumentContentProvider, EventEmitter, workspace } from "vscode";
import { CLI_TENANT_CMD } from "../constants";
import { ServiceBase } from "./service-base";
import { GraphClient, execShellCmd } from "../clients/graph-client";

export class OrganizationService extends ServiceBase {

    // The constructor for the OwnerService class.
    constructor(graphClient: GraphClient) {
        super(graphClient);
    }

    // Shows the tenant information.
    async showTenantInformation(): Promise<void> {

        const status = this.triggerTreeChange("Loading Tenant Information");

        // Execute the az cli command to get the tenant id
        const tenantId = await execShellCmd(CLI_TENANT_CMD);

        // Get the tenant information.
        const tenant = await this.graphClient.getTenantInformation(tenantId);

        const tenantInformation = {
            id: tenant.id,
            displayName: tenant.displayName,
            primaryDomain: tenant.verifiedDomains?.filter(x => x.isDefault === true)[0].name,
            verifiedDomains: tenant.verifiedDomains?.map(x => x.name)
        };

        const newDocument = new class implements TextDocumentContentProvider {
            onDidChangeEmitter = new EventEmitter<Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(): string {
                return JSON.stringify(tenantInformation, null, 4);
            }
        };

        this.disposable.push(workspace.registerTextDocumentContentProvider('tenantInformation', newDocument));
        const uri = Uri.parse('tenantInformation:' + tenant.displayName + ".json");
        workspace.openTextDocument(uri)
            .then(doc => {
                window.showTextDocument(doc, { preview: false });
                status!.dispose();
            });
    }
}
