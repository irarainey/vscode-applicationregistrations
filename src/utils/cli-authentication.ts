import { window } from "vscode";
import { execShellCmd } from "../utils/exec-shell-cmd";

// Invokes the Azure CLI sign-in command to authenticate the user.
export const authenticate = async (): Promise<boolean | undefined> => {
    // Prompt the user for the tenant name or Id.
    const tenant = await window.showInputBox({
        placeHolder: "Tenant Name or ID",
        prompt: "Enter the tenant name or ID, or leave blank for the default tenant",
        title: "Azure CLI Sign-In Tenant",
        ignoreFocusOut: true
    });

    // If the tenant is undefined then we don't want to do anything because they pressed cancel.
    if (tenant === undefined) {
        return undefined;
    }

    // Build the command to invoke the Azure CLI sign-in command.
    let command = "az login";
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
};