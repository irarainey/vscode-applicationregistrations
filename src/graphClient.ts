import "isomorphic-fetch";
import { scope, propertiesToIgnoreOnUpdate } from './constants';
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application } from "@microsoft/microsoft-graph-types";

// This is the client for the Microsoft Graph API
export class GraphClient {

    // A private instance of the graph client
    private client?: Client;

    // A public function triggered to determine the authentication state
    public authenticationStateChange: (state: boolean | undefined) => void;

    // Constructor
    constructor() {
        this.authenticationStateChange = () => { };
    }

    // Initialises the graph client
    public initialise(): void {

        // Create an Azure CLI credential
        const credential = new AzureCliCredential();

        // Create a TokenCredentialAuthenticationProvider with the Azure CLI credential and required scope
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: [scope] });

        // Create the graph client options
        const clientOptions: ClientOptions = {
            defaultVersion: "v1.0",
            debugLogging: false,
            authProvider
        };

        // Initialise the graph client with the defined options
        this.client = Client.initWithMiddleware(clientOptions);

        // Attempt to get an access token to determine the authentication state
        try {
            credential.getToken(scope)
                .then((_response) => {
                    // If the access token is returned, the user is authenticated
                    this.authenticationStateChange(true);
                })
                .catch((_error) => {
                    // If the access token is not returned, the user is not authenticated
                    this.authenticationStateChange(false);
                });
        } catch (error) {
            console.log("Error: " + error);
        }
    }

    // Returns to get the application registrations
    public async getApplicationsAll(filter?: string): Promise<Application[]> {
        const request = await this.client!.api("/applications/")
            .filter(filter === undefined ? "" : filter)
            .get();
        return request.value;
    }

    // Deletes an application registration
    public async deleteApplication(id: string): Promise<void> {
        await this.client!.api(`/applications/${id}`)
            .delete();
    }

    // Creates a new application registration
    public async createApplication(application: Application): Promise<Application> {
        return await this.client!.api("/applications/")
            .post(application);
    }

    // Updates an application registration
    public async updateApplication(id: string, application: Application): Promise<Application> {

        // Remove the properties that cannot be updated
        propertiesToIgnoreOnUpdate.forEach(property => {
            delete application[property as keyof Application];
        });

        return await this.client!.api(`/applications/${id}`)
            .update(application);
    }
}
