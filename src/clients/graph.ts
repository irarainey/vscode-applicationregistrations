import "isomorphic-fetch";
import * as ChildProcess from "child_process";
import { window, Disposable, workspace } from 'vscode';
import { scope, propertiesToIgnoreOnUpdate, directoryObjectsUri } from '../constants';
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application, PasswordCredential, ServicePrincipal } from "@microsoft/microsoft-graph-types";

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
            debugLogging: false,
            authProvider
        };

        // Initialise the graph client with the defined options
        this._client = Client.initWithMiddleware(clientOptions);

        // Attempt to get an access token to determine the authentication state
        await credential.getToken(scope)
            .then(() => {
                // If the access token is returned, the user is authenticated
                this._graphClientInitialised = true;
                this.initialiseTreeView("APPLICATIONS", window.setStatusBarMessage("$(loading~spin) Loading application registrations..."), undefined);
            })
            .catch(() => {
                // If the access token is not returned, the user is not authenticated
                this._graphClientInitialised = false;
                this.initialiseTreeView("SIGN-IN", undefined, undefined);
            });
    }

    // Invokes the Azure CLI sign-in command.
    public async cliSignIn(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        window.showInputBox({
            placeHolder: "Tenant name or Id...",
            prompt: "Enter the tenant name or Id, or leave blank for the default tenant",
            ignoreFocusOut: true
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
                this.execShellCmd(command)
                    .then(() => {
                        this.initialiseTreeView("INITIALISING", undefined, undefined);
                        this.initialise();
                    }).catch(() => {
                        this.initialise();
                    });
            });
    }

    // Returns a count of all owned application registrations
    public async getApplicationCountOwned(): Promise<number> {
        return await this._client!.api("/me/ownedObjects/$/Microsoft.Graph.Application/$count")
            .header("ConsistencyLevel", "eventual")
            .get();
    }

    // Returns a count of all application registrations
    public async getApplicationCountAll(): Promise<number> {
        return await this._client!.api("/applications/$count")
            .header("ConsistencyLevel", "eventual")
            .get();
    }

    // Returns full details for a specified application registration
    public async getApplicationDetailsFull(id: string): Promise<Application> {
        return await this._client!.api(`/applications/${id}`)
            .top(1)
            .get();
    }

    // Returns partial details for a specified application registration
    public async getApplicationDetailsPartial(id: string, select: string, expandOwners: boolean = false): Promise<Application> {
        if (expandOwners !== true) {
            return await this._client!.api(`/applications/${id}`)
                .top(1)
                .select(select)
                .get();
        } else {
            return await this._client!.api(`/applications/${id}`)
                .top(1)
                .select(select)
                .expand("owners")
                .get();
        }
    }

    // Returns ids and names for all owned application registrations
    public async getApplicationListOwned(filter?: string): Promise<Application[]> {
        const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
        if (useEventualConsistency === true) {
            const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
            const ownedApplications = await this._client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
                .filter(filter === undefined ? "" : filter)
                .header("ConsistencyLevel", "eventual")
                .count(true)
                .top(maximumApplicationsShown)
                .orderby("displayName")
                .select("id,displayName")
                .get();
            return ownedApplications.value;
        } else {
            const maximumQueryApps = workspace.getConfiguration("applicationregistrations").get("maximumQueryApps") as number;
            const ownedApplications = await this._client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
                .top(maximumQueryApps)
                .select("id,displayName")
                .get();
            return ownedApplications.value;
        }
    }

    // Returns ids and names for all application registrations
    public async getApplicationListAll(filter?: string): Promise<Application[]> {
        const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
        if (useEventualConsistency === true) {
            const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
            const allApplications = await this._client!.api("/applications/")
                .filter(filter === undefined ? "" : filter)
                .header("ConsistencyLevel", "eventual")
                .count(true)
                .top(maximumApplicationsShown)
                .orderby("displayName")
                .select("id,displayName")
                .get();
            return allApplications.value;
        } else {
            const maximumQueryApps = workspace.getConfiguration("applicationregistrations").get("maximumQueryApps") as number;
            const allApplications = await this._client!.api("/applications/")
                .top(maximumQueryApps)
                .select("id,displayName")
                .get();
            return allApplications.value;
        }
    }

    // Returns all owners for a specified application registration
    public async getApplicationOwners(id: string): Promise<any> {
        return await this._client!.api(`/applications/${id}/owners`)
            .get();
    }

    // Removes an owner from an application registration
    public async removeApplicationOwner(id: string, userId: string): Promise<void> {
        await this._client!.api(`/applications/${id}/owners/${userId}/$ref`)
            .delete();
    }

    // Find users by display name
    public async findUserByName(name: string): Promise<any> {
        return await this._client!.api("/users")
            .filter(`startswith(displayName, '${name}')`)
            .get();
    }

    // Find users by email address
    public async findUserByEmail(name: string): Promise<any> {
        return await this._client!.api("/users")
            .filter(`startswith(mail, '${name}')`)
            .get();
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
        return await this._client!.api(`/applications/${id}/addPassword`)
            .post({
                "passwordCredential": {
                    "endDateTime": expiry,
                    "displayName": description
                }
            });
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

    // Gets a service principal by application registration id
    public async getServicePrincipalByAppId(id: string): Promise<ServicePrincipal> {
        return await this._client!.api(`servicePrincipals(appId='${id}')`)
            .get();
    }

    // Gets a list of service principals by display name
    public async getServicePrincipalByDisplayName(name?: string): Promise<ServicePrincipal[]> {
        if (name === undefined || name === "") {
            const sp = await this._client!.api("servicePrincipals")
                .select("appId,appDisplayName,appDescription")
                .get();
            return sp.value;
        } else {
            const sp = await this._client!.api("servicePrincipals")
                .filter(`startswith(displayName, '${name}')`)
                .select("appId,appDisplayName,appDescription")
                .get();
            return sp.value;
        }
    }

    // Disposes the client
    public dispose(): void {
        this._client = undefined;
        this._graphClientInitialised = false;
    }

    private execShellCmd(cmd: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            ChildProcess.exec(cmd, (error: any, response: any) => {
                if (error) {
                    return reject(error);
                }
                return resolve(response);
            });
        });
    }
}