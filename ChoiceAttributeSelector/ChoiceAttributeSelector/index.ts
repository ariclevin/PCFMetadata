import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { ChoiceAttributeSelectorComponent, IChoiceAttributeSelectorProps } from "./ChoiceAttributeSelectorComponent";

export class ChoiceAttributeSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentOptionSet: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentOptionSet = context.parameters.Attribute.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IChoiceAttributeSelectorProps = {
            selectedAttribute: context.parameters.Attribute.raw || "",
            entity: context.parameters.Entity.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(ChoiceAttributeSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            Attribute: this.currentOptionSet
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedOptionSet: string): void {
        this.currentOptionSet = selectedOptionSet;
        this.notifyOutputChanged();
    }
}
