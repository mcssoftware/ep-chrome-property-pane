import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';

export interface IPropertyFieldEpChromeData {
    isActive: boolean;
    title?: string;
    showTitle: boolean;
    showIcon: boolean;
    iconPath?: string;
}

export interface IPropertyFieldEpChromeProps {
    key: string;
    /**
     * Label for the Chrome field.
     */
    label?: string;
    /**
     * Value to be displayed in the chrome
     */
    value?: IPropertyFieldEpChromeData;
    /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and an error message is displayed below the chrome field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
    onGetErrorMessage?: (value: IPropertyFieldEpChromeData) => string | Promise<string>;
}

/**
* Internal properties of PropertyFieldEpChrome custom field
*/
export interface IPropertyFieldEpChromePropsInternal extends IPropertyPaneCustomFieldProps, IPropertyFieldEpChromeProps {
}