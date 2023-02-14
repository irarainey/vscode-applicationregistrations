import { Application, Organization, User, RoleAssignment, PasswordCredential } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../../types/graph-result";
import { mockApplications, mockOrganizations, mockUser, mockRoleAssignments, mockUsers, mockNewPasswordKeyId } from "../../tests/test-data";

export class GraphApiRepository {
	async getApplicationCountOwned(): Promise<GraphResult<number>> {
		return { success: true, value: mockApplications.length };
	}

	async getApplicationListOwned(_filter?: string): Promise<GraphResult<Application[]>> {
		return { success: true, value: mockApplications };
	}

	async getApplicationDetailsFull(id: string): Promise<GraphResult<Application>> {
		return {
			success: true,
			value: mockApplications.filter((a) => a.id === id)[0]
		};
	}

	async getApplicationDetailsPartial(id: string, _select: string, _expandOwners: boolean = false): Promise<GraphResult<Application>> {
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

	async createApplication(application: Application): Promise<GraphResult<Application>> {
		return { success: true, value: { displayName: application.displayName } };
	}

	async addApplicationOwner(id: string, userId: string): Promise<GraphResult<void>> {
		const app: Application = mockApplications.filter((a) => a.id === id)[0];
		app.owners!.push(mockUsers.filter((u) => u.id === userId)[0]);
		return { success: true };
	}

	async getApplicationOwners(id: string): Promise<GraphResult<User[]>> {
		return { success: true, value: mockApplications.filter((a) => a.id === id)[0].owners };
	}

	async removeApplicationOwner(id: string, userId: string): Promise<GraphResult<void>> {
		const app: Application = mockApplications.filter((a) => a.id === id)[0];
		app.owners!.splice(app.owners!.findIndex((o) => o.id === userId), 1);
		return { success: true };
	}

	async deleteApplication(id: string): Promise<GraphResult<void>> {
		mockApplications.splice(mockApplications.findIndex((a) => a.id === id), 1);
		return { success: true };
	}

	async getTenantInformation(tenantId: string): Promise<GraphResult<Organization>> {
		return { success: true, value: mockOrganizations.filter((o) => o.id === tenantId)[0] };
	}

	async getUserInformation(): Promise<GraphResult<User>> {
		return { success: true, value: mockUser as User };
	}

	async getRoleAssignments(id: string): Promise<GraphResult<RoleAssignment[]>> {
		return { success: true, value: mockRoleAssignments.filter((r) => r.principalId === id) };
	}

	async getSignInAudience(id: string): Promise<GraphResult<string>> {
		return { success: true, value: mockApplications.filter((a) => a.id === id)[0].signInAudience };
	}

	async addPasswordCredential(id: string, description: string, expiry: string): Promise<GraphResult<PasswordCredential>> {
		const passwordCredential: PasswordCredential = {
			keyId: mockNewPasswordKeyId,
			displayName: description,
			customKeyIdentifier: null,
			endDateTime: expiry,
			hint: "newPassword",
			secretText: "NEWPASSWORD",
			startDateTime: "2023-01-01T00:00:00"
		};
		const app: Application = mockApplications.filter((a) => a.id === id)[0];
		app.passwordCredentials!.push(passwordCredential);
		return { success: true, value: passwordCredential };
	}

	async deletePasswordCredential(id: string, passwordId: string): Promise<GraphResult<void>> {
		const app: Application = mockApplications.filter((a) => a.id === id)[0];
		app.passwordCredentials!.splice(app.passwordCredentials!.findIndex((o) => o.keyId === passwordId), 1);
		return { success: true };
	}
}
