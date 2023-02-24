import { ServiceBase } from "./service-base";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";
import { Organization, User, RoleAssignment } from "@microsoft/microsoft-graph-types";
import { AccountProvider } from "../types/account-provider";
import { showJsonDocument } from "../utils/json-document-utils";

export class OrganizationService extends ServiceBase {

    // The constructor for the OrganizationService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider, private accountProvider: AccountProvider) {
        super(graphRepository, treeDataProvider);
    }

	// Shows the tenant information.
	async showTenantInformation(): Promise<void> {
		// Set the status bar message.
		const status = this.indicateChange("Loading Tenant Information...");

        await this.accountProvider.getAccountInformation()
            .then(async (response) => {
                await this.showTenantWindow(response.tenantId, status);
            })
            .catch(async (error) => {
                await this.handleError(error);
            });
    }

    // Shows the tenant information in a new read-only window.
    private async showTenantWindow(tenantId: string, status: string | undefined): Promise<void> {
        // Get the user information
        const user: GraphResult<User> = await this.graphRepository.getUserInformation();
        if (user.success !== true || user.value === undefined) {
            await this.handleError(user.error);
            return;
        }

		// Get the assigned directory roles.
		const roles: GraphResult<RoleAssignment[]> = await this.graphRepository.getRoleAssignments(user.value.id!);
		if (roles.success !== true || roles.value === undefined) {
			await this.handleError(roles.error);
			return;
		}

		const userInformation = {
			id: user.value.id,
			displayName: user.value.displayName,
			mail: user.value.mail,
			userPrincipalName: user.value.userPrincipalName,
			userType: user.value.userType,
			roles: roles.value.map((role) => {
				return {
					id: role.roleDefinition!.id,
					displayName: role.roleDefinition!.displayName,
					description: role.roleDefinition!.description
				};
			})
		};

		// Get the tenant information.
		const result: GraphResult<Organization> = await this.graphRepository.getTenantInformation(tenantId);
		if (result.success === true && result.value !== undefined) {
			const tenantInformation = {
				id: result.value.id,
				displayName: result.value.displayName,
				primaryDomain: result.value.verifiedDomains?.filter((x) => x.isDefault === true)[0].name,
				verifiedDomains: result.value.verifiedDomains?.map((x) => x.name),
				user: userInformation
			};
			this.disposable.push(await showJsonDocument(`Tenant - ${result.value.displayName!}`, tenantInformation, status));
		} else {
			await this.handleError(result.error);
			return;
		}
	}
}
