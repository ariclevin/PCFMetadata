import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { ViewSelectorComponent, IViewSelectorProps } from "./ViewSelectorComponent";

export class ViewSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentView: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentView = context.parameters.View.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IViewSelectorProps = {
            selectedView: context.parameters.View.raw || "",
            entity: context.parameters.Entity.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(ViewSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            View: this.currentView
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedView: string): void {
        this.currentView = selectedView;
        this.notifyOutputChanged();
    }
}
