import "isomorphic-fetch";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { DefaultAzureCredential } from "@azure/identity";
import { Application } from "@microsoft/microsoft-graph-types";

export class GraphClient {
    
    private client: Client;

    constructor() {
        const credential = new DefaultAzureCredential();
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["https://graph.microsoft.com/.default"] });
        const clientOptions: ClientOptions = {
            defaultVersion: "v1.0",
            debugLogging: false,
            authProvider
        };
        this.client = Client.initWithMiddleware(clientOptions);
    }

    public async getApplicationsAll(filter?: string): Promise<Application[]> {
        const request = await this.client.api("/applications/")
            .filter(filter === undefined ? "" : filter)
            .get();
        return request.value;
    }

    public async deleteApplication(id: string): Promise<void> {
        await this.client.api(`/applications/${id}`)
            .delete();
    }

    public async createApplication(application: Application): Promise<Application> {
        return await this.client.api("/applications/")
            .post(application);
    }

    public async updateApplication(id: string, application: Application): Promise<Application> {

        delete application['appId'];
        delete application['publisherDomain'];
        delete application['passwordCredentials'];
        delete application['deletedDateTime'];
        delete application['disabledByMicrosoftStatus'];
        delete application['createdDateTime'];
        delete application['serviceManagementReference'];
        delete application['parentalControlSettings'];
        delete application['keyCredentials'];
        delete application['groupMembershipClaims'];
        delete application['tags'];
        delete application['tokenEncryptionKeyId'];

        return await this.client.api(`/applications/${id}`)
            .update(application);
    }
}
