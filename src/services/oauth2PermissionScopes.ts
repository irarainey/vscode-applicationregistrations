import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';

export class OAuth2PermissionScopeService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // The constructor for the OAuth2PermissionScopeService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

}