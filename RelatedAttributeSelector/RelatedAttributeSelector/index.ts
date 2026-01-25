import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { RelatedAttributeSelectorComponent, IRelatedAttributeSelectorProps } from "./RelatedAttributeSelectorComponent";

export class RelatedAttributeSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentRelated: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentRelated = context.parameters.Attribute.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IRelatedAttributeSelectorProps = {
            selectedAttribute: context.parameters.Attribute.raw || "",
            formEntity: context.parameters.FormEntity.raw || "",
            queryEntity: context.parameters.QueryEntity.raw || "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(RelatedAttributeSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            Attribute: this.currentRelated
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedRelated: string): void {
        this.currentRelated = selectedRelated;
        this.notifyOutputChanged();
    }
}
