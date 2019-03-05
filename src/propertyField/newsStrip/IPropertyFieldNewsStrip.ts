export interface IPropertyFieldNewsStripData {
    numberOfItemsToDisplay: number;
    showAuthor: boolean;
    showArticleDate: boolean;
    showRating: boolean;
    showSummary: boolean;
}

export const getNewsStripDefaultValues = (): IPropertyFieldNewsStripData => {
    return {
        numberOfItemsToDisplay: 4,
        showAuthor: true,
        showArticleDate: true,
        showRating: true,
        showSummary: true
    };
};

export interface IPropertyFieldNewsStripProps {
    key: string;
    /**
     * Label for the Chrome field.
     */
    label?: string;
    /**
     * Value to be displayed in the chrome
     */
    value?: IPropertyFieldNewsStripData;
    /**
     * Whether the property pane field is enabled or not.
     */
    disabled?: boolean;
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
    onGetErrorMessage?: (value: IPropertyFieldNewsStripData) => string | Promise<string>;
    /**
     * Defines a onPropertyChange function to raise when the selected value changed.
     * Normally this function must be always defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
     * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
     * Default value is 200.
     */
    deferredValidationTime?: number;
}

export interface IPropertyFieldNewsStripPropsInternal extends IPropertyFieldNewsStripProps {
    targetProperty: string;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
}