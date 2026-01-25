import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { EntityCollectionSelectorComponent, IEntityCollectionSelectorProps } from "./EntityCollectionSelectorComponent";

export class EntityCollectionSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private selectedEntities: string[];

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.selectedEntities = this.parseEntityString(context.parameters.EntityNames.raw || "");
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IEntityCollectionSelectorProps = {
            selectedEntities: this.parseEntityString(context.parameters.EntityNames.raw || ""),
            activitiesOnly: context.parameters.activitiesOnly.raw || "All",
            supportsActivities: context.parameters.supportsActivities.raw || "All",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(EntityCollectionSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            EntityNames: this.selectedEntities.join(";")
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(selectedEntities: string[]): void {
        this.selectedEntities = selectedEntities;
        this.notifyOutputChanged();
    }

    private parseEntityString(entityString: string): string[] {
        return entityString ? entityString.split(";").filter(e => e.trim()) : [];
    }
}
