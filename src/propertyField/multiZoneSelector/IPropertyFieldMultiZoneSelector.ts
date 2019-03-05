import { IPropertyFieldNewsSelectorData as IPropertyFieldSelectorData, getPropertyFieldNewsSelectorDefaultValue } from "../newsSelector";
import { ISPTermStorePickerService } from "../../services/ISPTermStorePickerService";
import { ISPService } from "../../services/ISPService";

export enum ZoneDataType {
    None = 0,
    Article,
    Video,
    Content
}

export function getZoneDataType(value: number) {
    if (value === ZoneDataType.Article) {
        return ZoneDataType.Article;
    }
    if (value === ZoneDataType.Video) {
        return ZoneDataType.Video;
    }
    return ZoneDataType.Content;
}

export interface IContentData {
    title: string;
    backgroundUrl: string;
    backgroundColor: string;
    targetUrl: string;
    iconUrl: string;
}

export interface IVideoData {
    url: string;
}

export interface IZoneData {
    type: ZoneDataType;
    data: IContentData | IVideoData | IPropertyFieldSelectorData;
}

export interface IPropertyPaneMultiZoneSelectorData extends Array<IZoneData> { }

/**
* Get default value for Content Zone 
* @returns {IContentData}
*/
export const getContentDataDefaultValue = (): IContentData => {
    return {
        title: "",
        backgroundColor: "#ffffff",
        backgroundUrl: "",
        targetUrl: "",
        iconUrl: "",
    };
};

/**
* Get default value for Video Zone
* @returns {IVideoData}
*/
export const getVideoDataDefaultValue = (): IVideoData => {
    return {
        url: ""
    };
};

/**
* Get default value for Article Zone
* @returns {IPropertyFieldSelectorData}
*/
export const getArticleDataDefaultValue = (): IPropertyFieldSelectorData => {
    return getPropertyFieldNewsSelectorDefaultValue();
};

/**
* Get default value for zone, if type is specified it gets default value for that type
*
* @param {ZoneDataType} [type]
* @returns {IZoneData}
*/
export const getZoneDefaultValue = (type?: ZoneDataType): IZoneData => {
    if (typeof type === "undefined") {
        type = ZoneDataType.Content;
    }
    let data: IContentData | IVideoData | IPropertyFieldSelectorData;
    switch (type) {
        case ZoneDataType.Article: data = getArticleDataDefaultValue(); break;
        case ZoneDataType.Video: data = getVideoDataDefaultValue(); break;
        default: data = getContentDataDefaultValue(); break;
    }
    return {
        type,
        data
    };
};

/**
* Get default value for all zones.
*
* @returns {IPropertyPaneMultiZoneSelectorData}
*/
export const getPropertyFieldMultiZoneNewsSelectorDefaultValue = (): IPropertyPaneMultiZoneSelectorData => {
    return [];
};

/**
 * Public properties of the PropertyFieldTermPicker custom field
 */
export interface IPropertyFieldMultiZoneSelectorProps {
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
     * Define number of zones
     */
    numberOfZones: number;
    /**
     * Defines the selected by default term sets.
     */
    value?: IPropertyPaneMultiZoneSelectorData;
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
    onGetErrorMessage?: (value: IPropertyPaneMultiZoneSelectorData) => string | Promise<string>;
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
export interface IPropertyFieldMultiZoneSelectorPropsInternal extends IPropertyFieldMultiZoneSelectorProps {
    termService: ISPTermStorePickerService;
    spService: ISPService;
    targetProperty: string;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
}