import { AppRegDataProvider } from '../data/applicationRegistration';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class KeyCredentialService extends ServiceBase {

    // The constructor for the KeyCredentialsService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        super(dataProvider, graphClient);
    }
}