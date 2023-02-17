// Basic information about the currently signed-in account
export class AccountInformation {
	tenantId: string; // The tenant ID
	homeTenantId: string; // The home tenant ID
	environmentName: string; // The current cloud type - AzureCloud, AzureChinaCloud, AzureUSGovernment, AzureGermanCloud
	subscriptionName: string; // The subscription name
	subscriptionId: string; // The subscription ID

	constructor(tenantId: string = "", homeTenantId: string = "", environmentName: string = "", subscriptionName: string = "", subscriptionId: string = "") {
		this.tenantId = tenantId;
		this.homeTenantId = homeTenantId;
		this.environmentName = environmentName;
		this.subscriptionName = subscriptionName;
		this.subscriptionId = subscriptionId;
	}

	deserialize(jsonTextInput: string): AccountInformation {
		var data = null;
		try {
			data = JSON.parse(jsonTextInput);
		} catch (e) {
			// There was a problem parsing the JSON text input.
			console.error(e);
		}
		this.tenantId = data?.tenantId ?? "";
		this.homeTenantId = data?.homeTenantId ?? "";
		this.environmentName = data?.environmentName ?? "";
		this.subscriptionName = data?.name ?? "";
		this.subscriptionId = data?.id ?? "";
		return this;
	}
}

// Note that available cloud types can be found with "az cloud list": https://learn.microsoft.com/cli/azure/manage-clouds-azure-cli#list-available-clouds
