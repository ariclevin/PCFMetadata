import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { AttributeCollectionSelectorComponent, IAttributeCollectionSelectorProps } from "./AttributeCollectionSelectorComponent";

export class AttributeCollectionSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private selectedAttributes: string[];

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.selectedAttributes = this.parseAttributeString(context.parameters.Attribute.raw || "");
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IAttributeCollectionSelectorProps = {
            selectedAttributes: this.parseAttributeString(context.parameters.Attribute.raw || ""),
            entityName: context.parameters.EntityName.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(AttributeCollectionSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            Attribute: this.selectedAttributes.join(";")
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedAttributes: string[]): void {
        this.selectedAttributes = selectedAttributes;
        this.notifyOutputChanged();
    }

    private parseAttributeString(attributeString: string): string[] {
        return attributeString ? attributeString.split(";").filter(a => a.trim()) : [];
    }
}
