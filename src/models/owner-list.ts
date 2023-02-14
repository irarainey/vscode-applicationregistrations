import { User } from "@microsoft/microsoft-graph-types";

export class OwnerList {
	// The list of users in the directory.
	users: User[] | undefined = undefined;
}
