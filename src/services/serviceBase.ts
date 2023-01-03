import { Event, EventEmitter, Disposable, ThemeIcon, window } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { ActivityStatus } from '../interfaces/activityStatus';
import { AppRegItem } from '../models/appRegItem';

export class ServiceBase {

    // A private instance of the GraphClient class.
    protected _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    protected _dataProvider: AppRegDataProvider;

    // A private instance of the EventEmitter class to handle error events.
    protected _onError: EventEmitter<ActivityStatus> = new EventEmitter<ActivityStatus>();

    // A private instance of the EventEmitter class to handle complete events.
    protected _onComplete: EventEmitter<ActivityStatus> = new EventEmitter<ActivityStatus>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityStatus> = this._onError.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityStatus> = this._onComplete.event;

    // The constructor for the OwnerService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        this._dataProvider = dataProvider;
        this._graphClient = graphClient;
    }

    // Initiates the visual change of the tree view
    protected indicateChange(statusBarMessage: string, item: AppRegItem): Disposable {
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();
        return window.setStatusBarMessage(`$(loading~spin) ${statusBarMessage}`);
    }
}