import {
    PXScreen,
    createSingle,
    graphInfo,
    PXView,
    PXFieldState,
    PXFieldOptions,
} from "client-controls";

@graphInfo({
    graphType: "PX.Survey.Ext.SurveyComponentMaint",
    primaryView: "SUComponent",
})
export class SU204003 extends PXScreen {
    SUComponent = createSingle(SurveyComponent);
}

export class SurveyComponent extends PXView {
    ComponentID: PXFieldState<PXFieldOptions.CommitChanges>;
    Description: PXFieldState;
    ComponentType: PXFieldState<PXFieldOptions.CommitChanges>;
    Active: PXFieldState;
    Body: PXFieldState;
}
