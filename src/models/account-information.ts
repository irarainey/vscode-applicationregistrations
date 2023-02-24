export class AccountInformation {
	constructor(public tenantId: string = "", public homeTenantId: string = "", public environmentName: string = "", public subscriptionName: string = "", public subscriptionId: string = "") {
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
