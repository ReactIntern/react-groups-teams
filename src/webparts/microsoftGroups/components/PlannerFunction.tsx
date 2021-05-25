import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { MSGraphClient } from "@microsoft/sp-http";

export class PlannerFunction {
    constructor(private Tenant: string, private graphClient: MSGraphClient = null) { }

    public async GetPlanner(groupId: string): Promise<string> {
        const plans = await this.graphClient
          .api(`/groups/${groupId}/planner/plans`)
          .get();
    
        if (plans.value.length > 0) {
          var PlanID;
    
          // Note: Groups can have more than one plan, this
          // just picks the last one for simplicity's sake
          plans.value.map((plan: PlannerPlan) => {
            PlanID = plan.id;
          });
    
          return `https://tasks.office.com/${this.Tenant}.com/EN-US/Home/Planner#/plantaskboard?groupId=${groupId}&planId=${PlanID}`;
        }
      }

}