import { IPropertyFieldNewsStripPropsInternal, IPropertyPaneNewsStripData } from "../IPropertyPaneNewsStrip";

/**
 * News strip component Props
 *
 * @export
 * @interface IPropertyFieldNewsStripHostProps
 * @extends {IPropertyFieldNewsStripPropsInternal}
 */
export interface IPropertyFieldNewsStripHostProps extends IPropertyFieldNewsStripPropsInternal {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyPaneNewsStripData) => void;
}

export interface IPropertyFieldNewsStripHostHost{
    value: IPropertyPaneNewsStripData;
    numberOfItemsText: string;
    errorMessage?: string;
}
