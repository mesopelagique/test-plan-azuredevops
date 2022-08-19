import { Control } from "VSS/Controls";
import { IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItem, WorkItemType, WorkItemExpand, WorkItemRelation} from "TFS/WorkItemTracking/Contracts"
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { getClient } from "TFS/WorkItemTracking/RestClient";
import { idField, witField, projectField, titleField, parentField } from "./fieldNames";

export class TestPlanControl extends Control<{}> {
    // data
    private wiId: number;
    private requirements: WorkItem[];
    private testPlans: { [reqId: number]: WorkItem[] };
    private types: Map<string, WorkItemType> = new Map<string, WorkItemType>();

    // transform relations to items with all data
    private async relationToWorkItems(relations: WorkItemRelation[], project: string, expand?: WorkItemExpand) {
        var items: WorkItem[] = []
        for (const relation of relations) {
            const relationId = Number(relation.url.split('/').pop()) // not very clean way to get id, and project could differ?
            const item: WorkItem = await getClient().getWorkItem(relationId, undefined, undefined, expand, project);
            if (item) {
                items.push(item)
            }
        }
        return items
    }
    
    // Get children of type requirement, feature must have `relations` filled
    private async getRequirements(feature: WorkItem, project: string) {
        const childRelations = feature.relations.filter(relation => relation.rel == "System.LinkTypes.Hierarchy-Forward");
    
        var requirements: WorkItem[] = await this.relationToWorkItems(childRelations, project, WorkItemExpand.Relations)
        return requirements.filter(function(requirement) {
            return requirement.fields["System.WorkItemType"] == "Requirement"
        })
        .sort(function(x,y) {
            return x.fields["System.Title"].localeCompare(y.fields["System.Title"], 'en', {numeric: true});
        });
    }
    
    // Get element that test the requirement, requirement must have `relations` filled
    private async getTestCases(requirement: WorkItem, project: string) {
        const testCases = requirement.relations.filter(relation => relation.rel == "Microsoft.VSTS.Common.TestedBy-Forward");
        return this.relationToWorkItems(testCases, project, undefined);
    }

    private async fillTypes(wi: WorkItem, project: string) {
        const type = wi.fields[witField];
        if (!this.types.has(type)) {
            this.types.set(type, await getClient().getWorkItemType(project, type));
        }
    }

    private async fillTestPlan(wiID: number, type: string, parentId: number, project: string) {
        if(type == "Requirement") {
            const requirement: WorkItem = await getClient().getWorkItem(wiID, undefined, undefined, WorkItemExpand.Relations, project);
            this.requirements.push(requirement);
            
            const testCases = await this.getTestCases(requirement, project);
            this.testPlans[requirement.id] = testCases;
        } else if(type == "Feature") {
            const feature: WorkItem = await getClient().getWorkItem(wiID, undefined, undefined, WorkItemExpand.Relations, project);
            const requirements: WorkItem[] = await this.getRequirements(feature, project);
            this.requirements = requirements

            for (const requirement of requirements) {
                const testCases = await this.getTestCases(requirement, project);
                this.testPlans[requirement.id] = testCases;
            }

        } else {
            const item: WorkItem = await getClient().getWorkItem(parentId, [witField, parentField, titleField, idField], undefined, undefined, project);
            if (item && item.fields[parentField]!=undefined) {
                await this.fillTestPlan(item.id, item.fields[witField] as string, item.fields[parentField] as number, project);
            }
        }
    }
 
    public async refresh() {
        const formService = await WorkItemFormService.getService();
        const fields = await formService.getFieldValues([idField, projectField, parentField, witField]);
        this.wiId = fields[idField] as number;
        const project = fields[projectField] as string;
        this.requirements = [];
        this.testPlans = {};

        await this.fillTestPlan(this.wiId, fields[witField] as string, fields[parentField] as number, project);
        // update ui
        if (this.requirements && this.requirements.length != 0) {
            await this.updateTestPlan(project);
        } else {
            this.updateNoTestPlan();
        }
    }

    private updateNoTestPlan() {
        this._element.html(`<div class="no-tests-message">No test case</div>`);
        VSS.resize(window.innerWidth, $(".test-plan-callout").outerHeight() + 16)
    }

    private async updateTestPlan(project: string) {
        this._element.html("");
        const list = $("<div class=\"la-list\"></div>").appendTo(this._element);
   
        for (const requirement of this.requirements) {
            const item = $("<div class=\"la-item\"></div>").appendTo(list);
            const wrapper = $("<div class=\"la-item-wrapper\"></div>").appendTo(item);
            const artifactdata = $("<div class=\"la-artifact-data\"></div>").appendTo(wrapper);
            const primarydata = $("<div class=\"la-primary-data\"></div>").appendTo(artifactdata);
 
            const primaryicon = $("<div class=\"la-primary-icon\" style=\"display: inline;\">&nbsp;</div>").appendTo(primarydata);

            var type: WorkItemType = this.types.get(requirement.fields[witField])
            if (type == null) {
                await this.fillTypes(requirement, project);
                type = this.types.get(requirement.fields[witField])
            }
            if (type != null) {
                const iconsymbol = "bowtie-symbol-"+type.icon.id
                .replace("icon_", "")
                .replace("symbol-check_box", "status-success-box")
                .replace("_", ""); // no info how to convert api info and bowtie map
                const iconcolor = "#"+type.color;
                $("<span aria-hidden=\"true\" class=\"bowtie-icon "+iconsymbol+" flex-noshrink\" style=\"color: "+iconcolor+";\"> </span>&nbsp;").appendTo(primaryicon);
            }

            $("<div class=\"la-primary-data-id\" style=\"display: inline;\">"+requirement.id.toString()+"&nbsp;</div>").appendTo(primarydata);

            const link = $("<div class=\"ms-TooltipHost \" style=\"display: inline;\"></div>").appendTo(primarydata);
            $("<a/>").text(requirement.fields[titleField])
            .attr({
                href: requirement._links["html"]["href"],
                target: "_blank",
                title: "Navigate to item"
            }).appendTo(link);

            const testCases = this.testPlans[requirement.id];
            if (testCases) {
                for (const testCase of testCases) {
                    const item = $("<div class=\"la-item\"></div>").appendTo(list);
                    const wrapper = $("<div class=\"la-item-wrapper\"></div>").appendTo(item);
                    const artifactdata = $("<div class=\"la-artifact-data\"></div>").appendTo(wrapper);
                    const primarydata = $("<div class=\"la-primary-data\"></div>").appendTo(artifactdata);
         
                    const primaryicon = $("<div class=\"la-primary-icon\" style=\"display: inline;\">&nbsp;</div>").appendTo(primarydata);
        
                    var type: WorkItemType = this.types.get(testCase.fields[witField])
                    if (type == null) {
                        await this.fillTypes(testCase, project);
                        type = this.types.get(testCase.fields[witField])
                    }
                    if (type != null) {
                        const iconcolor = "#"+type.color;
                        $("<span aria-hidden=\"true\" class=\"bowtie-icon bowtie-test-case flex-noshrink\" style=\"color: "+iconcolor+";\"> </span>&nbsp;").appendTo(primaryicon);
                    }

                    $("<div class=\"la-primary-data-id\" style=\"display: inline;\">"+testCase.id.toString()+"&nbsp;</div>").appendTo(primarydata);
        
                    const link = $("<div class=\"ms-TooltipHost \" style=\"display: inline;\"></div>").appendTo(primarydata);
                    $("<a/>").text(testCase.fields[titleField])
                    .attr({
                        href: testCase._links["html"]["href"],
                        target: "_blank",
                        title: "Navigate to item"
                    }).appendTo(link);
                }
            }
        }
        VSS.resize();
    }

    public onLoaded(loadedArgs: IWorkItemLoadedArgs) {
        if (loadedArgs.isNew) {
            this._element.html(`<div class="new-wi-message">Save the work item to see the test plan data</div>`);
        } else {
            this.wiId = loadedArgs.id;
            this._element.html("");
            this._element.append($("<div/>").text("Looking for tests..."));
            this.refresh();
        }
    }

    public onRefreshed() {
        this.refresh();
    }

    public onSaved(_: IWorkItemChangedArgs) {
        this.refresh();
    }
}
