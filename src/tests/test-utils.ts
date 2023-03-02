import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";

export const getTopLevelTreeItem = async (objectId: string, treeDataProvider: AppRegTreeDataProvider, contextValue?: string): Promise<AppRegItem | undefined> => {
	let element: AppRegItem | undefined;

	switch (contextValue) {
		case undefined:
		case "LOGOUT-URL-PARENT":
			element = undefined;
			break;
		case "OWNERS":
		case "APPID-URIS":
		case "PASSWORD-CREDENTIALS":
		case "CERTIFICATE-CREDENTIALS":
		case "WEB-REDIRECT":
		case "SPA-REDIRECT":
		case "NATIVE-REDIRECT":
		case "APP-ROLES":
		case "EXPOSED-API-PERMISSIONS":
		case "API-PERMISSIONS":
			element = { objectId, contextValue };
			break;
		default:
			element = undefined;
			break;
	}

	// Get the children of the specified element. If element is undefined it will return the whole tree
	const tree = await treeDataProvider.getChildren(element);

	if (element === undefined) {
		// If the element is undefined then find the actual application we are looking for
		const app = tree!.find((x) => x.objectId === objectId);

		// Return the children of the context element that we are looking for. If the context element is undefined then return the application
		return contextValue === undefined ? app : app?.children?.find((x) => x.contextValue === contextValue);
	} else {
		// If an element is specified then return the element with the children
		return { ...element, children: toAppRegArray(tree) };
	}
};

const toAppRegArray = (value: AppRegItem[] | null | undefined): AppRegItem[] | undefined => {
	if (!value) {
		return undefined;
	}
	return value;
};
