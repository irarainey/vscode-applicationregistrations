import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class KeyCredentialService extends ServiceBase {

    // The constructor for the KeyCredentialsService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }
}