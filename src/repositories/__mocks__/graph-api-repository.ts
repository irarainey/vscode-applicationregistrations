import { GraphResult } from "../../types/graph-result";
import { Application } from "@microsoft/microsoft-graph-types";

export class GraphApiRepository {
  constructor() {

  }

  async updateApplication(id: string, application: Application): Promise<GraphResult<void>> {
    return { success: true };
  }

}
