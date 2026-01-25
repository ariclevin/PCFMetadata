import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FormSelectorComponent, IFormSelectorProps } from "./FormSelectorComponent";

export class FormSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentForm: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentForm = context.parameters.Form.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IFormSelectorProps = {
            selectedForm: context.parameters.Form.raw || "",
            entity: context.parameters.Entity.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(FormSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            Form: this.currentForm
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedForm: string): void {
        this.currentForm = selectedForm;
        this.notifyOutputChanged();
    }
}
