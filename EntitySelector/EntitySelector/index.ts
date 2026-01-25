import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { EntitySelectorComponent, IEntitySelectorProps } from "./EntitySelectorComponent";

export class EntitySelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentLogicalName: string;
    private currentDisplayName: string;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentLogicalName = context.parameters.logicalName.raw || "";
        this.currentDisplayName = context.parameters.displayName.raw || "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IEntitySelectorProps = {
            selectedLogicalName: context.parameters.logicalName.raw || "",
            selectedDisplayName: context.parameters.displayName.raw || "",
            activitiesOnly: context.parameters.activitiesOnly.raw || "All",
            supportsActivities: context.parameters.supportsActivities.raw || "All",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(EntitySelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            logicalName: this.currentLogicalName,
            displayName: this.currentDisplayName
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(logicalName: string, displayName: string): void {
        this.currentLogicalName = logicalName;
        this.currentDisplayName = displayName;
        this.notifyOutputChanged();
    }
}
