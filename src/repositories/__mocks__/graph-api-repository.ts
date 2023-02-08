import { Application, Organization, User, RoleAssignment } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../../types/graph-result";
import { mockApplications, mockOrganizations, mockUser, mockRoleAssignments } from "./mock-graph-data";

export class GraphApiRepository {
	async getApplicationCountOwned(): Promise<GraphResult<number>> {
		return { success: true, value: mockApplications.length };
	}

	async getApplicationListOwned(filter?: string): Promise<GraphResult<Application[]>> {
		return { success: true, value: mockApplications };
	}

	async getApplicationDetailsFull(id: string): Promise<GraphResult<Application>> {
		return {
			success: true,
			value: mockApplications.filter((a) => a.id === id)[0]
		};
	}

	async getApplicationDetailsPartial(id: string, select: string, expandOwners: boolean = false): Promise<GraphResult<Application>> {
		return {
			success: true,
			value: mockApplications.filter((a) => a.id === id)[0]
		};
	}

	async updateApplication(id: string, appChange: Application): Promise<GraphResult<void>> {
		const app = mockApplications.filter((a) => a.id === id)[0];

		if (appChange.web !== undefined) {
			app.web = { ...app.web, ...appChange.web };
		} else if (appChange.spa !== undefined) {
			app.spa = { ...app.spa, ...appChange.spa };
		} else if (appChange.api !== undefined) {
			app.api = { ...app.api, ...appChange.api };
		} else if (appChange.publicClient !== undefined) {
			app.publicClient = { ...app.publicClient, ...appChange.publicClient };
		} else {
			Object.assign(app, appChange);
		}

		return { success: true };
	}

	async getTenantInformation(tenantId: string): Promise<GraphResult<Organization>> {
		return { success: true, value: mockOrganizations.filter((o) => o.id === tenantId)[0] };
	}

	async getUserInformation(): Promise<GraphResult<User>> {
		return { success: true, value: mockUser };
	}

	async getRoleAssignments(id: string): Promise<GraphResult<RoleAssignment[]>> {
		return { success: true, value: mockRoleAssignments.filter((r) => r.principalId === id) };
	}

	async getSignInAudience(id: string): Promise<GraphResult<string>> {
		return { success: true, value: mockApplications.filter((a) => a.id === id)[0].signInAudience };
	}
}
