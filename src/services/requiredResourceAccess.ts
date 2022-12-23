import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';

export class RequiredResourceAccessService {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
    }

}