import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { ServiceBase } from "./service-base";
import { GraphClient } from "../clients/graph-client";

export class KeyCredentialService extends ServiceBase {

    // The constructor for the KeyCredentialsService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }
}