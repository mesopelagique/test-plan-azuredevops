import { Control } from "VSS/Controls";
import { IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItem, WorkItemType, WorkItemExpand, WorkItemRelation} from "TFS/WorkItemTracking/Contracts"
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { getClient } from "TFS/WorkItemTracking/RestClient";
import { idField, witField, projectField, titleField, parentField, testSteps } from "./fieldNames";

class TestStep {
    testCase: WorkItem
    index: number
    action: string
    expectedResult: string
    constructor(testCase: WorkItem, index: number, action: string, expectedResult: string) {
        this.testCase = testCase;
        this.index = index;
        this.action = action.replace(/<[^>]*>/g, "");
        this.expectedResult = expectedResult.replace(/<[^>]*>/g, "");
    }
}

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

    private getTestSteps(testCase: WorkItem) {
        const xmlString = testCase.fields[testSteps];

        const domParser = new DOMParser();
        const xmlDocument = domParser.parseFromString(xmlString, "text/xml");
        const steps = xmlDocument.querySelectorAll("step");
        var result: TestStep[] = [];
        var index = 1;
        for (const step of steps) {
            if (step.tagName == "step") {
                console.log(step.childNodes[0].textContent);
                console.log(step.childNodes[1].textContent);
                result.push(new TestStep(testCase, index, step.childNodes[0].textContent, step.childNodes[1].textContent));
                index += 1;
            }
        }
        return result;
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

    private appendNode(parent: any, id: number, icon: string, iconColor: string, text: string, href: string, level: number) {
        const item = $("<div class=\"la-item\" style=\"padding-left: "+(level * 12).toString()+"px;\"></div>").appendTo(parent);
        const wrapper = $("<div class=\"la-item-wrapper\"></div>").appendTo(item);
        const artifactdata = $("<div class=\"la-artifact-data\"></div>").appendTo(wrapper);
        const primarydata = $("<div class=\"la-primary-data\"></div>").appendTo(artifactdata);

        const primaryicon = $("<div class=\"la-primary-icon\" style=\"display: inline;\">&nbsp;</div>").appendTo(primarydata);

        if (iconColor.length>0) {
            $("<span aria-hidden=\"true\" class=\"bowtie-icon "+icon+" flex-noshrink\";\" style=\"color: "+iconColor+";\"> </span>&nbsp;").appendTo(primaryicon);
        } else {
            $("<span aria-hidden=\"true\" class=\"bowtie-icon "+icon+" flex-noshrink\";\"> </span>&nbsp;").appendTo(primaryicon);
        }
        if (id>=0) {
            $("<div class=\"la-primary-data-id\" style=\"display: inline;\">&nbsp;"+id.toString()+"&nbsp;</div>").appendTo(primarydata);
        } else {
            $("<div class=\"la-primary-data-id\" style=\"display: inline;\">&nbsp;&nbsp;</div>").appendTo(primarydata);
        }
        const link = $("<div class=\"ms-TooltipHost \" style=\"display: inline;\"></div>").appendTo(primarydata);
        $("<a/>").text(text)
        .attr({
            href: href,
            target: "_blank",
            title: "Navigate to item"
        }).appendTo(link);
    }

    private async getWorkItemType(item: WorkItem, project: string) {
        var type: WorkItemType = this.types.get(item.fields[witField])
        if (type == null) {
            await this.fillTypes(item, project);
            type = this.types.get(item.fields[witField])
        }
        return type
    }

    private async updateTestPlan(project: string) {
        this._element.html("");
        const list = $("<div class=\"la-list\"></div>").appendTo(this._element);
   
        for (const requirement of this.requirements) {
            var iconsymbol = "bowtie-symbol-stickynote"
            var iconcolor = ""
            var type: WorkItemType = await this.getWorkItemType(requirement, project);
            if (type != null) {
                iconsymbol = "bowtie-symbol-"+type.icon.id
                .replace("icon_", "")
                .replace("symbol-check_box", "status-success-box")
                .replace("_", ""); // no info how to convert api info and bowtie map
                iconcolor = "#"+type.color;
            }
            this.appendNode(list, requirement.id, iconsymbol, iconcolor, requirement.fields[titleField], requirement._links["html"]["href"], 0);

            const testCases = this.testPlans[requirement.id];
            if (testCases) {
                for (const testCase of testCases) {
                    $("<div><hr style=\"margin: 0 0 0 12px; height: 5px; border: 0px solid #D6D6D6; border-top-width: 1px;\">").appendTo(list);
                    iconcolor = ""
                    type = await this.getWorkItemType(testCase, project);
                    if (type != null) {
                        iconcolor = "#"+type.color;
                    }
                    this.appendNode(list, testCase.id, "bowtie-test-case", iconcolor, testCase.fields[titleField], testCase._links["html"]["href"], 1);
                    const testSteps = this.getTestSteps(testCase);
                    for (const testStep of testSteps) {
                        this.appendNode(list, testStep.index, "bowtie-step", "", testStep.action, testCase._links["html"]["href"], 2);
                        if (testStep.expectedResult.length>0) {
                            this.appendNode(list, -1, "bowtie-watch-eye", "", testStep.expectedResult, testCase._links["html"]["href"], 3);
                        }
                    }
                }
            }
            $("<div><hr/></div>").appendTo(list);
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
