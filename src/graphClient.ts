import "isomorphic-fetch";
import * as vscode from 'vscode';
import { view, scope } from './constants';
import { SignInDataProvider } from './dataProviders/signInDataProvider';
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application } from "@microsoft/microsoft-graph-types";

export class GraphClient {

    private client: Client;
    public token?: string;
    public authenticated: () => void;

    constructor() {

        this.authenticated = () => { };

        const credential = new AzureCliCredential();
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: [scope] });
        const clientOptions: ClientOptions = {
            defaultVersion: "v1.0",
            debugLogging: false,
            authProvider
        };

        this.client = Client.initWithMiddleware(clientOptions);

        try {
            credential.getToken(scope)
                .then((response) => {
                    this.token = response.token;
                    this.authenticated();
                })
                .catch((error) => {
                    this.token = undefined;
                    vscode.window.registerTreeDataProvider(view, new SignInDataProvider());
                    vscode.window.showInformationMessage("An authentication error occurred. Please ensure you are signed in to Azure CLI.");
                });
        } catch (error) {
            console.log("Error: " + error);
        }
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
