import { IPropertyFieldEpChromeProps, IPropertyFieldEpChromeData } from "../IPropertyFieldEpChrome";

export interface IPropertyPaneEpChromeHostProps extends IPropertyFieldEpChromeProps {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyFieldEpChromeData) => void;
}

export interface IPropertyPaneEpChromeHostState {
    value: IPropertyFieldEpChromeData;
    errorMessage?: string;
}