import { AppRegDataProvider } from '../data/applicationRegistration';
import { ServiceBase } from './serviceBase';

export class KeyCredentialService extends ServiceBase {

    // The constructor for the KeyCredentialsService class.
    constructor(dataProvider: AppRegDataProvider) {
        super(dataProvider);
    }
}