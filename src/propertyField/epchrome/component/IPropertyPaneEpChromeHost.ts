import { IPropertyFieldEpChromeProps, IPropertyFieldEpChromeData } from "../IPropertyFieldEpChrome";

export interface IPropertyPaneEpChromeHostProps extends IPropertyFieldEpChromeProps {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyFieldEpChromeData) => void;
}

export enum ColorPickerType{
    background=0,
    text,
}

export interface IPropertyPaneEpChromeHostState {
    value: IPropertyFieldEpChromeData;
    errorMessage?: string;
    showColorPanel: boolean;
    panelColor: string;
    panelType?: ColorPickerType;
}