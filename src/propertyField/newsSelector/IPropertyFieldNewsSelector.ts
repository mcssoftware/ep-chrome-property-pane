import { IPickerTerms } from "./termStoreEntity";
import { ISPTermStorePickerService } from "../../services/ISPTermStorePickerService";
import { ISPService } from "../../services/ISPService";

export enum ActiveDisplayModeType {
    Latest = 0,
    Specific
}

export interface IPropertyFieldNewsSelectorData {
    ActiveDisplayMode: ActiveDisplayModeType;
    NewsChannel: IPickerTerms;
    ArticleId: number;
}

export const getPropertyFieldDefaultValue = (): IPropertyFieldNewsSelectorData => {
    return {
        ActiveDisplayMode: ActiveDisplayModeType.Latest,
        ArticleId: 0,
        NewsChannel: []
    };
};

/**
 * Public properties of the PropertyFieldTermPicker custom field
 */
export interface IPropertyFieldNewsSelectorProps {
    /**
     * Property field label displayed on top
     */
    label: string;
    /**
     * TermSet Picker Panel title
     */
    panelTitle: string;
    /**
     * Defines if the user can select only one or many term sets. Default value is false.
     */
    allowMultipleSelections?: boolean;
    /**
     * Defines the selected by default term sets.
     */
    initialValues?: IPropertyFieldNewsSelectorData;
    /**
     * Indicator to define if the system Groups are exclude. Default is false.
     */
    excludeSystemGroup?: boolean;
    /**
     * WebPart's context
     */
    context: any;
    /**
     * Limit the term sets that can be used by the group name or ID
     */
    limitByGroupNameOrID?: string;
    /**
     * Limit the terms that can be picked by the Term Set name or ID
     */
    limitByTermsetNameOrID?: string;
    /**
     * Defines a onPropertyChange function to raise when the selected value changed.
     * Normally this function must be always defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
     * Parent Web Part properties
     */
    properties: any;
    /**
     * An UNIQUE key indicates the identity of this control
     */
    key: string;
    /**
     * Whether the property pane field is enabled or not.
     */
    disabled?: boolean;
    /**
     * The method is used to get the validation error message and determine whether the input value is valid or not.
     *
     *   When it returns string:
     *   - If valid, it returns empty string.
     *   - If invalid, it returns the error message string and the text field will
     *     show a red border and show an error message below the text field.
     *
     *   When it returns Promise<string>:
     *   - The resolved value is display as error message.
     *   - The rejected, the value is thrown away.
     *
     */
    onGetErrorMessage?: (value: IPropertyFieldNewsSelectorData) => string | Promise<string>;
    /**
     * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
     * Default value is 200.
     */
    deferredValidationTime?: number;
    /**
     * Specifies if you want to show or hide the term store name from the panel
     */
    hideTermStoreName?: boolean;
    /**
     * Specify if the term set itself is selectable in the tree view
     */
    isTermSetSelectable?: boolean;
    /**
     * Specify which terms should be disabled in the term set so that they cannot be selected
     */
    disabledTermIds?: string[];

    /**
     * The delay time in ms before resolving suggestions, which is kicked off when input has been changed. e.g. if a second input change happens within the resolveDelay time, the timer will start over. 
     * Only until after the timer completes will onResolveSuggestions be called.
     * Default is 500
     */
    resolveDelay?: number;
}

/**
* Private properties of the PropertyFieldTermPicker custom field.
* We separate public & private properties to include onRender & onDispose method waited
* by the PropertyFieldCustom, witout asking to the developer to add it when he's using
* the PropertyFieldTermPicker.
*/
export interface IPropertyFieldNewsSelectorPropsInternal extends IPropertyFieldNewsSelectorProps {
    termService: ISPTermStorePickerService;
    spService: ISPService;
    targetProperty: string;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
}