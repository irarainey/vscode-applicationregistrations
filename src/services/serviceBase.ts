import { Event, EventEmitter, Disposable, ThemeIcon, window } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { ActivityResult } from '../interfaces/activityResult';
import { AppRegItem } from '../models/appRegItem';

export class ServiceBase {

    // A protected array of disposables for the service.
    protected _disposable: Disposable[] = [];

    // A protected instance of the GraphClient class.
    protected _graphClient: GraphClient;

    // A protected instance of the AppRegDataProvider class.
    protected _dataProvider: AppRegDataProvider;

    // A protected instance of the EventEmitter class to handle error events.
    protected _onError: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A protected instance of the EventEmitter class to handle complete events.
    protected _onComplete: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityResult> = this._onError.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityResult> = this._onComplete.event;

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

    // Dispose of anything that needs to be disposed of.
    public dispose(): void {
        this._onError.dispose();
        this._onComplete.dispose();
        while (this._disposable.length) {
            const x = this._disposable.pop();
            if (x) {
                x.dispose();
            }
        }
    }
}