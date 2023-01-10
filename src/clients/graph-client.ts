import "isomorphic-fetch";
import * as ChildProcess from "child_process";
import { SCOPE, PROPERTIES_TO_IGNORE_ON_UPDATE, DIRECTORY_OBJECTS_URI } from "../constants";
import { window, Disposable, workspace } from "vscode";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { AzureCliCredential } from "@azure/identity";
import { Application, KeyCredential, Organization, PasswordCredential, ServicePrincipal } from "@microsoft/microsoft-graph-types";

// This is the client for the Microsoft Graph API
export class GraphClient {

    // A private instance of the graph client
    private client?: Client;

    // A public property for the initialisation state
    public isGraphClientInitialised: boolean = false;

    // Constructor for the graph client
    constructor() {
        this.initialiseTreeView = () => { };
    }

    // A function that is called when the tree view is ready to be initialised
    initialiseTreeView: (type: string, statusBarMessage?: Disposable | undefined, filter?: string) => void;

    // Initialises the graph client
    async initialise(): Promise<void> {

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
        await credential.getToken(SCOPE)
            .then((result) => {
                // If the access token is returned, the user is authenticated
                console.log(result.token);
                this.isGraphClientInitialised = true;
                this.initialiseTreeView("APPLICATIONS", window.setStatusBarMessage("$(loading~spin) Loading Application Registrations"));
            })
            .catch(() => {
                // If the access token is not returned, the user is not authenticated
                this.isGraphClientInitialised = false;
                this.initialiseTreeView("SIGN-IN");
            });
    }

    // Invokes the Azure CLI sign-in command.
    async cliSignIn(): Promise<void> {
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
            }).catch(() => {
                this.initialise();
            });
    }

    // Returns a count of all owned application registrations
    async getApplicationCountOwned(): Promise<number> {
        return await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application/$count")
            .header("ConsistencyLevel", "eventual")
            .get();
    }

    // Returns a count of all application registrations
    async getApplicationCountAll(): Promise<number> {
        return await this.client!.api("/applications/$count")
            .header("ConsistencyLevel", "eventual")
            .get();
    }

    // Returns full details for a specified application registration
    async getApplicationDetailsFull(id: string): Promise<Application> {
        return await this.client!.api(`/applications/${id}`)
            .top(1)
            .get();
    }

    // Returns partial details for a specified application registration
    async getApplicationDetailsPartial(id: string, select: string, expandOwners: boolean = false): Promise<Application> {
        if (expandOwners !== true) {
            return await this.client!.api(`/applications/${id}`)
                .top(1)
                .select(select)
                .get();
        } else {
            return await this.client!.api(`/applications/${id}`)
                .top(1)
                .select(select)
                .expand("owners")
                .get();
        }
    }

    // Returns ids and names for all owned application registrations
    async getApplicationListOwned(filter?: string): Promise<Application[]> {
        const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
        if (useEventualConsistency === true) {
            const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
            const ownedApplications = await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
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
            const ownedApplications = await this.client!.api("/me/ownedObjects/$/Microsoft.Graph.Application")
                .top(maximumQueryApps)
                .select("id,displayName")
                .get();
            return ownedApplications.value;
        }
    }

    // Returns ids and names for all application registrations
    async getApplicationListAll(filter?: string): Promise<Application[]> {
        const useEventualConsistency = workspace.getConfiguration("applicationregistrations").get("useEventualConsistency") as boolean;
        if (useEventualConsistency === true) {
            const maximumApplicationsShown = workspace.getConfiguration("applicationregistrations").get("maximumApplicationsShown") as number;
            const allApplications = await this.client!.api("/applications/")
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
            const allApplications = await this.client!.api("/applications/")
                .top(maximumQueryApps)
                .select("id,displayName")
                .get();
            return allApplications.value;
        }
    }

    // Returns all owners for a specified application registration
    async getApplicationOwners(id: string): Promise<any> {
        return await this.client!.api(`/applications/${id}/owners`)
            .get();
    }

    // Removes an owner from an application registration
    async removeApplicationOwner(id: string, userId: string): Promise<void> {
        await this.client!.api(`/applications/${id}/owners/${userId}/$ref`)
            .delete();
    }

    // Find users by display name
    async findUsersByName(name: string): Promise<any> {
        return await this.client!.api("/users")
            .filter(`startswith(displayName, '${escapeSingleQuotesForFilter(name)}')`)
            .get();
    }

    // Find users by email address
    async findUsersByEmail(name: string): Promise<any> {
        return await this.client!.api("/users")
            .filter(`startswith(mail, '${escapeSingleQuotesForFilter(name)}')`)
            .get();
    }

    // Adds an owner to an application registration
    async addApplicationOwner(id: string, userId: string): Promise<void> {
        await this.client!.api(`/applications/${id}/owners/$ref`)
            .post({
                // eslint-disable-next-line @typescript-eslint/naming-convention
                "@odata.id": `${DIRECTORY_OBJECTS_URI}${userId}`
            });
    }

    // Adds a password credential to an application registration
    async addPasswordCredential(id: string, description: string, expiry: string): Promise<PasswordCredential> {
        return await this.client!.api(`/applications/${id}/addPassword`)
            .post({
                "passwordCredential": {
                    "endDateTime": expiry,
                    "displayName": description
                }
            });
    }

    // Deletes a password credential from an application registration
    async deletePasswordCredential(id: string, passwordId: string): Promise<void> {
        await this.client!.api(`/applications/${id}/removePassword`)
            .post({
                "keyId": passwordId
            });
    }

    // Updates the key credential collection from an application registration
    async updateKeyCredentials(id: string, credentials: KeyCredential[]): Promise<void> {
        await this.client!.api(`/applications/${id}`)
            .patch({
                "id": id,
                "keyCredentials": credentials
            });
    }

    // Creates a new application registration
    async createApplication(application: Application): Promise<Application> {
        return await this.client!.api("/applications/")
            .post(application);
    }

    // Updates an application registration
    async updateApplication(id: string, application: Application): Promise<Application> {
        // Remove the properties that cannot be updated
        PROPERTIES_TO_IGNORE_ON_UPDATE.forEach(property => {
            delete application[property as keyof Application];
        });
        return await this.client!.api(`/applications/${id}`)
            .update(application);
    }

    // Deletes an application registration
    async deleteApplication(id: string): Promise<void> {
        await this.client!.api(`/applications/${id}`)
            .delete();
    }

    // Gets a service principal by application registration id
    async findServicePrincipalByAppId(id: string): Promise<ServicePrincipal> {
        return await this.client!.api(`servicePrincipals(appId='${id}')`)
            .get();
    }

    // Gets a list of service principals by display name
    async findServicePrincipalsByDisplayName(name: string): Promise<ServicePrincipal[]> {
        const servicePrincipals = await this.client!.api("servicePrincipals")
            //.filter(`startswith(displayName, '${escapeSingleQuotes(name)}')`)
            .header("ConsistencyLevel", "eventual")
            .count(true)
            .search(`"displayName:${escapeSingleQuotesForSearch(name)}"`)
            .select("appId,appDisplayName,appDescription")
            .get();
        return servicePrincipals.value;
    }

    // Gets the tenant information.
    async getTenantInformation(tenantId: string): Promise<Organization> {
        // Get the tenant information
        return await this.client!.api(`/organization/${tenantId}`)
            .select("id,displayName,verifiedDomains")
            .get();
    }

    // Disposes the client
    dispose(): void {
        this.client = undefined;
        this.isGraphClientInitialised = false;
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