import "isomorphic-fetch";
import * as ChildProcess from "child_process";
import { SCOPE, PROPERTIES_TO_IGNORE_ON_UPDATE, DIRECTORY_OBJECTS_URI } from "../constants";
import { window, Disposable, workspace } from "vscode";
import { Client, ClientOptions, GraphRequest } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application, KeyCredential, User, PasswordCredential, ServicePrincipal, Organization } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";

// This is the client for the Microsoft Graph API
export class GraphApiRepository {

    // A private instance of the graph client
    private client?: Client;

    // A public property for the initialisation state
    public isClientInitialised: boolean = false;

    // Constructor for the graph client
    constructor() {
        this.initialiseTreeView = () => { };
    }

    // A function that is called when the tree view is ready to be initialised
    initialiseTreeView: (type: string, statusBarMessage?: Disposable | undefined, filter?: string) => void;

    // Initialises the graph client
    initialise() {

        // Create an Azure CLI credential
        const credential = new AzureCliCredential();

        // Create a TokenCredentialAuthenticationProvider with the Azure CLI credential and required scope
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: [SCOPE] });

        // Create the graph client options
        const clientOptions: ClientOptions = {
            defaultVersion: "v1.0",
            debugLogging: false,
            authProvider
        };

        // Initialise the graph client with the defined options
        this.client = Client.initWithMiddleware(clientOptions);

        // Attempt to get an access token to determine the authentication state
        credential.getToken(SCOPE)
            .then(() => {
                // If the access token is returned, the user is authenticated
                this.isClientInitialised = true;
                this.initialiseTreeView("APPLICATIONS", window.setStatusBarMessage("$(loading~spin) Loading Application Registrations"));
            })
            .catch(() => {
                // If the access token is not returned, the user is not authenticated
                this.isClientInitialised = false;
                this.initialiseTreeView("SIGN-IN");
            });
    }

    // Invokes the Azure CLI sign-in command to authenticate the user.
    async authenticate(): Promise<void> {
        // Prompt the user for the tenant name or Id.
        const tenant = await window.showInputBox({
            placeHolder: "Tenant Name or ID",
            prompt: "Enter the tenant name or ID, or leave blank for the default tenant",
            title: "Azure CLI Sign-In Tenant",
            ignoreFocusOut: true
        });

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
                this.initialiseTreeView("INITIALISING");
                this.initialise();
            })
            .catch(() => {
                this.initialise();
            });
    }

    // Returns a count of all owned application registrations
    async getApplicationCountOwned(): Promise<GraphResult<number>> {
        try {
            const result: number = await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application/$count")
                .header("ConsistencyLevel", "eventual")
                .get();
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns a count of all application registrations
    async getApplicationCountAll(): Promise<GraphResult<number>> {
        try {
            const result: number = await this.client!.api("/applications/$count")
                .header("ConsistencyLevel", "eventual")
                .get();
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns partial details for a specified application registration
    async getApplicationDetailsPartial(id: string, select: string, expandOwners: boolean = false): Promise<GraphResult<Application>> {
        try {
            if (expandOwners !== true) {
                const result: Application = await this.client!.api(`/applications/${id}`)
                    .top(1)
                    .select(select)
                    .get();
                return { success: true, value: result };
            } else {
                const result: Application = await this.client!.api(`/applications/${id}`)
                    .top(1)
                    .select(select)
                    .expand("owners")
                    .get();
                return { success: true, value: result };
            }
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns ids and names for all owned application registrations
    async getApplicationListOwned(filter?: string): Promise<GraphResult<Application[]>> {
        try {
            const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
            if (useEventualConsistency === true) {
                const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
                const result: any = await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
                    .filter(filter === undefined ? "" : filter)
                    .header("ConsistencyLevel", "eventual")
                    .count(true)
                    .top(maximumApplicationsShown)
                    .orderby("displayName")
                    .select("id,displayName")
                    .get();
                return { success: true, value: result.value };
            } else {
                const maximumQueryApps = workspace.getConfiguration("applicationregistrations").get("maximumQueryApps") as number;
                const result: any = await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
                    .top(maximumQueryApps)
                    .select("id,displayName")
                    .get();
                return { success: true, value: result.value };
            }
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns ids and names for all application registrations
    async getApplicationListAll(filter?: string): Promise<GraphResult<Application[]>> {
        try {
            const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
            if (useEventualConsistency === true) {
                const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
                const result: any = await this.client!.api("/applications/")
                    .filter(filter === undefined ? "" : filter)
                    .header("ConsistencyLevel", "eventual")
                    .count(true)
                    .top(maximumApplicationsShown)
                    .orderby("displayName")
                    .select("id,displayName")
                    .get();
                return { success: true, value: result.value };
            } else {
                const maximumQueryApps = workspace.getConfiguration("applicationregistrations").get("maximumQueryApps") as number;
                const result: any = await this.client!.api("/applications/")
                    .top(maximumQueryApps)
                    .select("id,displayName")
                    .get();
                return { success: true, value: result.value };
            }
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns all owners for a specified application registration
    async getApplicationOwners(id: string): Promise<GraphResult<User[]>> {
        try {
            const result: any = await this.client!.api(`/applications/${id}/owners`)
                .get();
            return { success: true, value: result.value };
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Returns full details for a specified application registration
    async getApplicationDetailsFull(id: string): Promise<GraphResult<Application>> {
        try {
            const result: Application = await this.client!.api(`/applications/${id}`)
                .top(1)
                .get();
            return { success: true, value: result };
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Removes an owner from an application registration
    async removeApplicationOwner(id: string, userId: string): Promise<GraphResult<void>> {
        try {
            return await this.client!.api(`/applications/${id}/owners/${userId}/$ref`)
                .delete()
                .then(() => {
                    return { success: true };
                });
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Adds an owner to an application registration
    async addApplicationOwner(id: string, userId: string): Promise<GraphResult<void>> {
        try {
            return await this.client!.api(`/applications/${id}/owners/$ref`)
                .post({
                    // eslint-disable-next-line @typescript-eslint/naming-convention
                    "@odata.id": `${DIRECTORY_OBJECTS_URI}${userId}`
                })
                .then(() => {
                    return { success: true };
                });
        }
        catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Adds a password credential to an application registration
    async addPasswordCredential(id: string, description: string, expiry: string): Promise<GraphResult<PasswordCredential>> {
        try {
            const result: PasswordCredential = await this.client!.api(`/applications/${id}/addPassword`)
                .post({
                    "passwordCredential": {
                        "endDateTime": expiry,
                        "displayName": description
                    }
                });
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Deletes a password credential from an application registration
    async deletePasswordCredential(id: string, passwordId: string): Promise<GraphResult<void>> {
        try {
            return await this.client!.api(`/applications/${id}/removePassword`)
                .post({
                    "keyId": passwordId
                })
                .then(() => {
                    return { success: true };
                });
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Updates the key credential collection from an application registration
    async updateKeyCredentials(id: string, credentials: KeyCredential[]): Promise<GraphResult<void>> {
        try {
            return await this.client!.api(`/applications/${id}`)
                .patch({
                    "id": id,
                    "keyCredentials": credentials
                })
                .then(() => {
                    return { success: true };
                });
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Creates a new application registration
    async createApplication(application: Application): Promise<GraphResult<Application>> {
        try {
            const result: Application = await this.client!.api("/applications/")
                .post(application);
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Updates an application registration
    async updateApplication(id: string, application: Application): Promise<GraphResult<void>> {
        try {
            // Remove the properties that cannot be updated
            PROPERTIES_TO_IGNORE_ON_UPDATE.forEach(property => {
                delete application[property as keyof Application];
            });
            return await this.client!.api(`/applications/${id}`)
                .update(application)
                .then(() => {
                    return { success: true };
                });
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Deletes an application registration
    async deleteApplication(id: string): Promise<GraphResult<void>> {
        try {
            return await this.client!.api(`/applications/${id}`)
                .delete()
                .then(() => {
                    return { success: true };
                });
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Find users by display name
    async findUsersByName(name: string): Promise<GraphResult<User[]>> {
        try {
            const result: any = await this.client!.api("/users")
                .filter(`startswith(displayName, '${escapeSingleQuotesForFilter(name)}')`)
                .get();
            return { success: true, value: result.value };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Find users by email address
    async findUsersByEmail(name: string): Promise<GraphResult<User[]>> {
        try {
            const result: any = await this.client!.api("/users")
                .filter(`startswith(mail, '${escapeSingleQuotesForFilter(name)}')`)
                .get();
            return { success: true, value: result.value };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Gets a service principal by application registration id
    async findServicePrincipalByAppId(id: string): Promise<GraphResult<ServicePrincipal>> {
        try {
            const result: ServicePrincipal = await this.client!.api(`servicePrincipals(appId='${id}')`)
                .get();
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Gets a list of service principals by display name
    async findServicePrincipalsByDisplayName(name: string): Promise<GraphResult<ServicePrincipal[]>> {
        try {
            const result: any = await this.client!.api("servicePrincipals")
                .header("ConsistencyLevel", "eventual")
                .count(true)
                .search(`"displayName:${escapeSingleQuotesForSearch(name)}"`)
                .select("appId,appDisplayName,appDescription")
                .get();
            return { success: true, value: result.value };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Gets the tenant information.
    async getTenantInformation(tenantId: string): Promise<GraphResult<Organization>> {
        try {
            // Get the tenant information
            const result: Organization = await this.client!.api(`/organization/${tenantId}`)
                .select("id,displayName,verifiedDomains")
                .get();
            return { success: true, value: result };
        } catch (error: any) {
            return { success: false, error: error };
        }
    }

    // Disposes the client
    dispose(): void {
        this.client = undefined;
        this.isClientInitialised = false;
    }
}

export const escapeSingleQuotesForFilter = (str: string) => {
    return str.replace(/'/g, "''");
};

export const escapeSingleQuotesForSearch = (str: string) => {
    return str.replace(/'/g, "\'");
};

export const execShellCmd = async (cmd: string) => {
    return new Promise<string>((resolve, reject) => {
        ChildProcess.exec(cmd, (error: any, response: any) => {
            if (error) {
                return reject(error);
            }
            return resolve(response);
        });
    });
};