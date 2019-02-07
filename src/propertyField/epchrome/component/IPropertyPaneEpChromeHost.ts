import { IPropertyFieldEpChromeProps, IPropertyFieldEpChromeData } from "../IPropertyFieldEpChrome";

export interface IPropertyPaneEpChromeHostProps extends IPropertyFieldEpChromeProps {
    /**
       * Callback for the onChanged event.
       */
    onChanged?: (newValue: IPropertyFieldEpChromeData) => void;
}

export interface IPropertyPaneEpChromeHostState {
    value: IPropertyFieldEpChromeData;
}