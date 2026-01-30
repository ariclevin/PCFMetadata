import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ModelDrivenAppSelectorComponent, IModelDrivenAppSelectorProps } from "./ModelDrivenAppSelectorComponent";
import * as React from "react";

export class ModelDrivenAppSelector implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private currentAppId: string;
    private currentAppName: string;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.currentAppId = context.parameters.appId.raw ?? "";
        this.currentAppName = context.parameters.appName.raw ?? "";
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        const props: IModelDrivenAppSelectorProps = {
            selectedAppId: context.parameters.appId.raw ?? "",
            selectedAppName: context.parameters.appName.raw ?? "",
            disabled: context.mode.isControlDisabled,
            onChange: this.onChange.bind(this)
        };

        return React.createElement(ModelDrivenAppSelectorComponent, props);
    }

    public getOutputs(): IOutputs {
        return {
            appId: this.currentAppId,
            appName: this.currentAppName
        };
    }

    public destroy(): void {
        // Cleanup if needed
    }

    private onChange(appId: string, appName: string): void {
        this.currentAppId = appId;
        this.currentAppName = appName;
        this.notifyOutputChanged();
    }
}
