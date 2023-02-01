import { GraphApiRepository } from "../../repositories/graph-api-repository";
//import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, ConfigurationTarget } from "vscode";
import { AppRegItem } from "../../models/app-reg-item";

export class AppRegTreeDataProvider {

  // This is the event that is fired when the tree view is refreshed.
  //private onDidChangeTreeDataEvent: EventEmitter<AppRegItem | undefined | null | void> = new EventEmitter<AppRegItem | undefined | null | void>();

  constructor(graphRepository: GraphApiRepository) {

  }

  triggerOnDidChangeTreeData(item?: AppRegItem) {
    //this.onDidChangeTreeDataEvent.fire(item);
}

}
