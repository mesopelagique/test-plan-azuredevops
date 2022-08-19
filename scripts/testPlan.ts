/// <reference types="vss-web-extension-sdk" />

import { TestPlanControl } from "./testPlanControl";
import * as Controls from "VSS/Controls";
import { IWorkItemNotificationListener } from "TFS/WorkItemTracking/ExtensionContracts";
 
const control = <TestPlanControl>Controls.BaseControl.createIn(TestPlanControl, $(".test-plan-control"));

const contextData: Partial<IWorkItemNotificationListener> = {
    onSaved: (savedEventArgs) => control.onSaved(savedEventArgs),
    onRefreshed: () => control.onRefreshed(),
    onLoaded: (loadedArgs) => control.onLoaded(loadedArgs)
};

VSS.register(VSS.getContribution().id, contextData);