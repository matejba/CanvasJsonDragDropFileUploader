/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    dropText: ComponentFramework.PropertyTypes.StringProperty;
    maxFiles: ComponentFramework.PropertyTypes.WholeNumberProperty;
    allowedFormats: ComponentFramework.PropertyTypes.StringProperty;
    reset: ComponentFramework.PropertyTypes.WholeNumberProperty;
    saveButtonColor: ComponentFramework.PropertyTypes.StringProperty;
    filesJson: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    dropText?: string;
    maxFiles?: number;
    allowedFormats?: string;
    reset?: number;
    saveButtonColor?: string;
    filesJson?: string;
}
