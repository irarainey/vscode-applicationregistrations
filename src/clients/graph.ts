import "isomorphic-fetch";
import { workspace, window, Disposable } from 'vscode';
import { scope, propertiesToIgnoreOnUpdate, directoryObjectsUri } from '../constants';
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application, PasswordCredential, User } from "@microsoft/microsoft-graph-types";
import { execShellCmd } from "../utils/shellUtils";

// This is the client for the Microsoft Graph API
export class GraphClient {

    // A private instance of the graph client
    private _client?: Client;

    // A private instance of the initialisation state
    private _graphClientInitialised: boolean = false;

    // A function that is called when the tree view is ready to be initialised
    public initialiseTreeView: (type: string, statusBarMessage: Disposable | undefined, filter?: string) => void;

    // Constructor for the graph client
    constructor() {
        this.initialiseTreeView = () => { };
    }

    // A public getter for the initialisation state
    public get isGraphClientInitialised(): boolean {
        return this._graphClientInitialised;
    }

    // A public setter for the initialisation state
    public set isGraphClientInitialised(value: boolean) {
        this._graphClientInitialised = value;
    }

    // Initialises the graph client
    public async initialise(): Promise<void> {

        // Create an Azure CLI credential
        const credential = new AzureCliCredential();

        // Create a TokenCredentialAuthenticationProvider with the Azure CLI credential and required scope
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: [scope] });

        // Create the graph client options
        const clientOptions: ClientOptions = {
            defaultVersion: "v1.0",
            debugLogging: true,
            authProvider
        };

        // Initialise the graph client with the defined options
        this._client = Client.initWithMiddleware(clientOptions);

        // Attempt to get an access token to determine the authentication state
        await credential.getToken(scope)
            .then((response) => {
                // If the access token is returned, the user is authenticated
                this._graphClientInitialised = true;
                //console.log(response.token);
                this.initialiseTreeView("APPLICATIONS", window.setStatusBarMessage("$(loading~spin) Loading Application Registrations..."), undefined);
            })
            .catch((error) => {
                // If the access token is not returned, the user is not authenticated
                this._graphClientInitialised = false;
                //console.error(error);
                this.initialiseTreeView("SIGN-IN", undefined, undefined);
            });
    }

    // Invokes the Azure CLI sign-in command.
    public async cliSignIn(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        window.showInputBox({
            placeHolder: "Tenant name or Id...",
            prompt: "Enter the tenant name or Id, or leave blank for the default tenant",
        })
            .then((tenant) => {
                // If the tenant is undefined then we don't want to do anything because they pressed cancel.
                if (tenant === undefined) {
                    return;
                }

                // Build the command to invoke the Azure CLI sign-in command.
                let command = "az login";
                if (tenant.length > 0) {
                    command += ` --tenant ${tenant}`;
                }

                // Execute the command.
                execShellCmd(command)
                    .then(() => {
                        this.initialise();
                    }).catch(() => {
                        this.initialise();
                    });
            });
    }

    // Returns all application registrations
    public async getApplicationsAll(filter?: string): Promise<Application[]> {

        const maximumReturned = workspace.getConfiguration("applicationregistrations")
            .get("maximumApplicationsReturned") as number;

        const request = await this._client!.api("/applications/")
            .header("ConsistencyLevel", "eventual")
            .count(true)
            .orderby("displayName")
            .filter(filter === undefined ? "" : filter)
            .top(maximumReturned)
            .get();
        return request.value;
    }

    // Returns owned application registrations
    public async getApplicationsOwned(filter?: string): Promise<Application[]> {

        const maximumReturned = workspace.getConfiguration("applicationregistrations")
            .get("maximumApplicationsReturned") as number;

        const request = await this._client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
            .header("ConsistencyLevel", "eventual")
            .count(true)
            .orderby("displayName")
            .filter(filter === undefined ? "" : filter)
            .top(maximumReturned)
            .get();
        return request.value;
    }

    // Returns all owners for a specified application registration
    public async getApplicationOwners(id: string): Promise<User[]> {
        const request = await this._client!.api(`/applications/${id}/owners`)
            .get();
        return request.value;
    }

    // Removes an owner from an application registration
    public async removeApplicationOwner(id: string, userId: string): Promise<void> {
        await this._client!.api(`/applications/${id}/owners/${userId}/$ref`)
            .delete();
    }

    // Find users by display name
    public async findUserByName(name: string): Promise<User[]> {
        const request = await this._client!.api("/users")
            .filter(`startswith(displayName, '${name}')`)
            .get();
        return request.value;
    }

    // Find users by email address
    public async findUserByEmail(name: string): Promise<User[]> {
        const request = await this._client!.api("/users")
            .filter(`startswith(mail, '${name}')`)
            .get();
        return request.value;
    }

    // Adds an owner to an application registration
    public async addApplicationOwner(id: string, userId: string): Promise<void> {
        await this._client!.api(`/applications/${id}/owners/$ref`)
            .post({
                // eslint-disable-next-line @typescript-eslint/naming-convention
                "@odata.id": `${directoryObjectsUri}${userId}`
            });
    }

    // Deletes a password credential from an application registration
    public async addPasswordCredential(id: string, description: string, expiry: string): Promise<PasswordCredential> {
        const request = await this._client!.api(`/applications/${id}/addPassword`)
            .post({
                "passwordCredential": {
                    "endDateTime": expiry,
                    "displayName": description
                }
            });
        return request;
    }

    // Deletes a password credential from an application registration
    public async deletePasswordCredential(id: string, passwordId: string): Promise<void> {
        await this._client!.api(`/applications/${id}/removePassword`)
            .post({
                "keyId": passwordId
            });
    }

    // Deletes an application registration
    public async deleteApplication(id: string): Promise<void> {
        await this._client!.api(`/applications/${id}`)
            .delete();
    }

    // Creates a new application registration
    public async createApplication(application: Application): Promise<Application> {
        return await this._client!.api("/applications/")
            .post(application);
    }

    // Updates an application registration
    public async updateApplication(id: string, application: Application): Promise<Application> {
        // Remove the properties that cannot be updated
        propertiesToIgnoreOnUpdate.forEach(property => {
            delete application[property as keyof Application];
        });

        return await this._client!.api(`/applications/${id}`)
            .update(application);
    }
}
