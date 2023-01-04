import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class RequiredResourceAccessService extends ServiceBase {

    // The constructor for the RequiredResourceAccessService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }
}