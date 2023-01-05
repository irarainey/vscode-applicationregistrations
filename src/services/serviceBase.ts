import { Event, EventEmitter, Disposable, ThemeIcon, window, Uri } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { ActivityResult } from '../interfaces/activityResult';
import { AppRegItem } from '../models/appRegItem';

export class ServiceBase {

    // A protected array of disposables for the service.
    private _disposable: Disposable[] = [];

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegTreeDataProvider class.
    private _treeDataProvider: AppRegTreeDataProvider;

    // A protected instance of the EventEmitter class to handle error events.
    private _onError: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A protected instance of the EventEmitter class to handle complete events.
    private _onComplete: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityResult> = this._onError.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityResult> = this._onComplete.event;

    // The constructor for the OwnerService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        this._treeDataProvider = treeDataProvider;
        this._graphClient = graphClient;
    }

    // A protected property to expose the disposable array.
    protected get disposable(): Disposable[] {
        return this._disposable;
    }

    // A protected property to expose the graph client.
    protected get graphClient(): GraphClient {
        return this._graphClient;
    }

    // A protected property to expose the data provider.
    protected get dataProvider(): AppRegTreeDataProvider {
        return this._treeDataProvider;
    }

    // Trigger the event to indicate an error
    protected triggerOnError(item: ActivityResult) {
        this._onError.fire(item);
    }

    // Trigger the event to indicate completion
    protected triggerOnComplete(item: ActivityResult) {
        this._onComplete.fire(item);
    }

    // Initiates the visual change of the tree view
    protected triggerTreeChange(statusBarMessage?: string, item?: AppRegItem): Disposable | undefined {
        if (item !== undefined) {
            item.iconPath = new ThemeIcon("loading~spin");
            this._treeDataProvider.triggerOnDidChangeTreeData(item);
        }

        if (statusBarMessage !== undefined) {
            return window.setStatusBarMessage(`$(loading~spin) ${statusBarMessage}`);
        }
    }

    // Sets the icon for a tree item
    protected setTreeItemIcon(item: AppRegItem, icon?: string | Uri | {
        light: string | Uri;
        dark: string | Uri;
    } | ThemeIcon | undefined, spinner?: boolean) {

        if (spinner) {
            icon = new ThemeIcon("loading~spin");
        }
        item.iconPath = icon;
        this._treeDataProvider.triggerOnDidChangeTreeData(item);
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