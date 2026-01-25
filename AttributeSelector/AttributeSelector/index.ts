import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { AttributeSelectorComponent, IAttributeSelectorProps } from "./AttributeSelectorComponent";

export class AttributeSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentAttributeName: string;
    private currentEntityName: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentAttributeName = context.parameters.attributeName.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IAttributeSelectorProps = {
            selectedAttribute: context.parameters.attributeName.raw || "",
            entity: context.parameters.entityName.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(AttributeSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            attributeName: this.currentAttributeName
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(attributeName: string): void {
        this.currentAttributeName = attributeName;
        this.notifyOutputChanged();
    }
}
