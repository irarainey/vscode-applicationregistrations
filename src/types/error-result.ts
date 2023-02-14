import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";

// This is the type for the return result of service methods
export type ErrorResult = {
	error?: Error;
	item?: AppRegItem;
	treeDataProvider?: AppRegTreeDataProvider;
};
