import { AppRegDataProvider } from '../data/applicationRegistration';
import { ServiceBase } from './serviceBase';

export class RequiredResourceAccessService extends ServiceBase {

    // The constructor for the RequiredResourceAccessService class.
    constructor(dataProvider: AppRegDataProvider) {
        super(dataProvider);
    }
}