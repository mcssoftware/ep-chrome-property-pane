import { IPropertyFieldNewsStripPropsInternal, IPropertyFieldNewsStripData } from "../IPropertyFieldNewsStrip";

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
    onChange: (targetProperty?: string, newValue?: IPropertyFieldNewsStripData) => void;
}

export interface IPropertyFieldNewsStripHostHost{
    value: IPropertyFieldNewsStripData;
    numberOfItemsText: string;
    errorMessage?: string;
}
