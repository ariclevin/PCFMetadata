import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { LookupAttributeSelectorComponent, ILookupAttributeSelectorProps } from "./LookupAttributeSelectorComponent";

export class LookupAttributeSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentLookup: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentLookup = context.parameters.Attribute.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: ILookupAttributeSelectorProps = {
            selectedAttribute: context.parameters.Attribute.raw || "",
            entity: context.parameters.Entity.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(LookupAttributeSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            Attribute: this.currentLookup
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedLookup: string): void {
        this.currentLookup = selectedLookup;
        this.notifyOutputChanged();
    }
}
