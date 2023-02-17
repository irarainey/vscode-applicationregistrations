import { ApiApplication, AppRole, User } from "@microsoft/microsoft-graph-types";
import { addYears, isAfter, isBefore, isDate } from "date-fns";
import { errorHandler } from "../error-handler";
import { OwnerList } from "../models/owner-list";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";

// Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
export const validateRedirectUri = (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined): string | undefined => {
	// Check to see if the uri already exists.
	if ((isEditing === true && oldValue !== uri) || isEditing === false) {
		if (existingRedirectUris.includes(uri)) {
			return "The Redirect URI specified already exists in this application.";
		}
	}

	if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
		// Check the redirect URI starts with https://
		if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
			return "The Redirect URI is not valid. A Redirect URI must start with https:// unless it is using http://localhost.";
		}
	} else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
		// Check the redirect URI starts with https:// or http:// or customScheme://
		if (uri.includes("://") === false) {
			return "The Redirect URI is not valid. A Redirect URI must start with https, http, or customScheme://.";
		}
	}

	// Check the length of the redirect URI.
	if (uri.length > 256) {
		return "The Redirect URI is not valid. A Redirect URI cannot be longer than 256 characters.";
	}

	return undefined;
};

// Validates the expiry date.
export const validatePasswordCredentialExpiryDate = (expiry: string): string | undefined => {
	const expiryDate = Date.parse(expiry);
	const now = Date.now();
	const maximumExpireDate = addYears(new Date(), 2);

	//Check if the expiry date is a valid date.
	if (Number.isNaN(expiryDate)) {
		return "Expiry must be a valid date.";
	}

	// Check if the expiry date is in the future.
	if (isBefore(expiryDate, now)) {
		return "Expiry must be in the future.";
	}

	// Check if the expiry date is less than 2 years in the future.
	if (isAfter(expiryDate, maximumExpireDate)) {
		return "Expiry must be less than 2 years in the future.";
	}

	return undefined;
};

// Validates the admin display name of an scope.
export const validateScopeAdminDisplayName = (displayName: string): string | undefined => {
	// Check the length of the display name.
	if (displayName.length > 100) {
		return "An admin display name cannot be longer than 100 characters.";
	}

	// Check the length of the display name.
	if (displayName.length < 1) {
		return "An admin display name cannot be empty.";
	}

	return undefined;
};

// Validates the user display name of an scope.
export const validateScopeUserDisplayName = (displayName: string): string | undefined => {
	// Check the length of the display name.
	if (displayName.length > 100) {
		return "An admin display name cannot be longer than 100 characters.";
	}

	return undefined;
};

// Validates the value of an scope.
export const validateScopeValue = async (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication): Promise<string | undefined> => {
	// Check the length of the value.
	switch (signInAudience) {
		case "AzureADMyOrg":
		case "AzureADMultipleOrgs":
			if (value.length > 120) {
				return "A value cannot be longer than 120 characters.";
			}
			break;
		case "AzureADandPersonalMicrosoftAccount":
		case "PersonalMicrosoftAccount":
			if (value.length > 40) {
				return "A value cannot be longer than 40 characters.";
			}
			break;
		default:
			break;
	}

	// Check the length of the value.
	if (value.length < 1) {
		return "A scope value cannot be empty.";
	}

	if (value.includes(" ")) {
		return "A scope value cannot contain spaces.";
	}

	if (value.startsWith(".")) {
		return "A scope value cannot start with a full stop.";
	}

	// Check to see if the value already exists.
	if (isEditing !== true || (isEditing === true && existingValue !== value)) {
		if (scopes.oauth2PermissionScopes!.find((r) => r.value === value) !== undefined) {
			return "The scope value specified already exists.";
		}
	}

	return undefined;
};

// Validates the value of the admin consent description.
export const validateScopeAdminDescription = (description: string): string | undefined => {
	// Check the length of the admin consent description.
	if (description.length < 1) {
		return "An admin consent description cannot be empty.";
	}

	return undefined;
};

// Validates the display name of an app role.
export const validateAppRoleDisplayName = (displayName: string): string | undefined => {
	// Check the length of the display name.
	if (displayName.length > 100) {
		return "A display name cannot be longer than 100 characters.";
	}

	// Check the length of the display name.
	if (displayName.length < 1) {
		return "A display name cannot be empty.";
	}

	return undefined;
};

// Validates the value of an app role.
export const validateAppRoleValue = (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]): string | undefined => {
	// Check the length of the value.
	if (value.length > 250) {
		return "A role value cannot be longer than 250 characters.";
	}

	// Check the length of the display name.
	if (value.length < 1) {
		return "A role value cannot be empty.";
	}

	if (value.includes(" ")) {
		return "A role value cannot contain spaces.";
	}

	if (value.startsWith(".")) {
		return "A role value cannot start with a full stop.";
	}

	// Check to see if the value already exists.
	if (isEditing !== true || (isEditing === true && oldValue !== value)) {
		if (roles!.find((r) => r.value === value) !== undefined) {
			return "The role value specified already exists.";
		}
	}

	return undefined;
};

// Validates the value of an app role.
export const validateAppRoleDescription = (description: string): string | undefined => {
	// Check the length of the description.
	if (description.length < 1) {
		return "A description cannot be empty.";
	}

	return undefined;
};

// Validates the display name of the application.
export const validateApplicationDisplayName = (name: string, signInAudience: string): string | undefined => {
	// Check the length of the application name.
	if (name.length < 1) {
		return "An application name must be at least one character.";
	}
	if ((signInAudience === "AzureADMyOrg" || signInAudience === "AzureADMultipleOrgs") && name.length > 120) {
		return "An application name cannot be longer than 120 characters.";
	} else if ((signInAudience === "AzureADandPersonalMicrosoftAccount" || signInAudience === "PersonalMicrosoftAccount") && name.length > 90) {
		return "An application name cannot be longer than 90 characters.";
	}

	return undefined;
};

// Validates the app id URI.
export const validateAppIdUri = (uri: string, signInAudience: string): string | undefined => {
	if (uri.endsWith("/") === true) {
		return "The Application ID URI cannot end with a trailing slash.";
	}

	if (signInAudience === "AzureADMyOrg" || signInAudience === "AzureADMultipleOrgs") {
		if (uri.includes("://") === false || uri.startsWith("://") === true) {
			return "The Application ID URI is not valid. It must start with http://, https://, api://, MS-APPX://, or customScheme://.";
		}
		if (uri.includes("*") === true) {
			return "Wildcards are not supported.";
		}
		if (uri.length > 255) {
			return "The Application ID URI is not valid. A URI cannot be longer than 255 characters.";
		}
	} else if (signInAudience === "AzureADandPersonalMicrosoftAccount" || signInAudience === "PersonalMicrosoftAccount") {
		if (uri.startsWith("api://") === false && uri.startsWith("https://") === false && uri.startsWith("http://") === false) {
			return "The Application ID URI is not valid. It must start with http://, https://, or api://.";
		}
		if (uri.includes("?") === true || uri.includes("#") === true || uri.includes("*") === true) {
			return "Wildcards, fragments, and query strings are not supported.";
		}
		if (uri.length > 120) {
			return "The Application ID URI is not valid. A URI cannot be longer than 120 characters.";
		}
	}

	return undefined;
};

// Validates the logout url.
export const validateLogoutUrl = (uri: string): string | undefined => {
	if (uri.startsWith("https://") === false) {
		return "The Logout URL is not valid. It must start with https://.";
	}
	if (uri.length > 256) {
		return "The Logout URL is not valid. A URL cannot be longer than 256 characters.";
	}

	return undefined;
};

// Validates the owner name or email address.
export const validateOwner = async (owner: string, existing: User[], graphRepository: GraphApiRepository, owners: OwnerList): Promise<string | undefined> => {
	// Check if the owner name is empty.
	if (owner === undefined || owner === null || owner.length === 0) {
		return undefined;
	}

	owners.users = undefined;
	let identifier: string = "";
	if (owner.indexOf("@") > -1) {
		// Try to find the user by email.
		const result: GraphResult<User[]> = await graphRepository.findUsersByEmail(owner);
		if (result.success === true && result.value !== undefined) {
			owners.users = result.value;
			identifier = "user with an email address";
		} else {
			await errorHandler({ error: result.error });
			return;
		}
	} else {
		// Try to find the user by name.
		const result: GraphResult<User[]> = await graphRepository.findUsersByName(owner);
		if (result.success === true && result.value !== undefined) {
			owners.users = result.value;
			identifier = "name";
		} else {
			await errorHandler({ error: result.error });
			return;
		}
	}

	if (owners.users.length === 0) {
		// User not found
		return `No ${identifier} beginning with ${owner} can be found in your directory.`;
	} else if (owners.users.length > 1) {
		// More than one user found
		return `More than one user with the ${identifier} beginning with ${owner} exists in your directory.`;
	}

	// Check if the user is already an owner.
	for (let i = 0; i < existing.length; i++) {
		if (existing[i].id === owners.users[0].id) {
			return `${owner} is already an owner of this application.`;
		}
	}

	return undefined;
};
