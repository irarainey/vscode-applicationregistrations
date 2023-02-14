import { AccountProvider } from "../types/account-provider";
import { AccountInformation } from "../models/account-information";
import { execShellCmd } from "./exec-shell-cmd";
import { CLI_LOGIN_CMD, CLI_LOGOUT_CMD, CLI_ACCOUNT_INFO_CMD } from "../constants";

// Provides functionality from the Azure CLI related to the currently signed-in account
export class AzureCliAccountProvider implements AccountProvider {
	// Invokes the Azure CLI sign-in command to sign the user in.
	async loginUser(tenant: string): Promise<boolean> {
		// Build the command to invoke the Azure CLI sign-in command.
		let command = CLI_LOGIN_CMD;
		if (tenant.length > 0) {
			command += ` --tenant ${tenant}`;
		}

		// Execute the command.
		const status = await execShellCmd(command)
			.then(() => {
				return true;
			})
			.catch(() => {
				return false;
			});
		return status;
	}

	// Invokes the Azure CLI sign-out command to sign the user out.
	async logoutUser(): Promise<void> {
		// Invokes the Azure CLI sign-out command to sign the user out.
		await execShellCmd(CLI_LOGOUT_CMD);
	}

	// Invokes the Azure CLI command to get the information about the current account (including tenant ID and cloud type)
	async getAccountInformation(): Promise<AccountInformation> {
		// Execute the Azure CLI command to get the information about the current account (including tenant ID and cloud type)
		const accountJsonText = await execShellCmd(CLI_ACCOUNT_INFO_CMD);
		return new AccountInformation().deserialize(accountJsonText);
	}
}
