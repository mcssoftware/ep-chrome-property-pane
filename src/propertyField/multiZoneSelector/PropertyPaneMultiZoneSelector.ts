import * as React from "react";
import * as ReactDom from "react-dom";
import { IPropertyFieldMultiZoneSelectorPropsInternal, IPropertyPaneMultiZoneSelectorData, getPropertyFieldMultiZoneNewsSelectorDefaultValue, IPropertyFieldMultiZoneSelectorProps } from "./IPropertyPaneMultiZoneSelector";
import { IPropertyPaneField, PropertyPaneFieldType, IWebPartContext } from "@microsoft/sp-webpart-base";
import { ISPTermStorePickerService } from "../../services/ISPTermStorePickerService";
import { ISPService } from "../../services/ISPService";
import SPTermStorePickerService from "../../services/SPTermStorePickerService";
import SPService from "../../services/SPService";
import { IPropertyFieldMultiZoneSelectorHostProps } from "./component/IPropertyFieldMultiZoneSelectorHost";
import { PropertyFieldMultiZoneNewsSelectorHost } from "./component/PropertyFieldMultiZoneSelectorHost";

/**
 * Represents a PropertyFieldTermPicker object.
 * NOTE: INTERNAL USE ONLY
 * @internal
 */
class PropertyPaneMultiZoneNewsSelectorBuilder implements IPropertyPaneField<IPropertyFieldMultiZoneSelectorPropsInternal>{
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyFieldMultiZoneSelectorPropsInternal;

    private label: string;
    private context: IWebPartContext;
    private key: string;
    private panelTitle: string;
    private limitByGroupNameOrID: string = null;
    private limitByTermsetNameOrID: string = null;
    private hideTermStoreName: boolean;
    private isTermSetSelectable: boolean;
    private disabledTermIds: string[];
    private termService: ISPTermStorePickerService;
    private spService: ISPService;
    private allowMultipleSelections: boolean = false;
    private excludeSystemGroup: boolean = false;
    private numberOfZones: number;
    private propertyValue: IPropertyPaneMultiZoneSelectorData = getPropertyFieldMultiZoneNewsSelectorDefaultValue();

    public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void { }
    private disabled: boolean = false;
    private onGetErrorMessage: (value: IPropertyPaneMultiZoneSelectorData) => string | Promise<string>;
    private deferredValidationTime: number = 200;

    /**
   * Constructor method
   */
    public constructor(_targetProperty: string, _properties: IPropertyFieldMultiZoneSelectorPropsInternal) {
        this.render = this.render.bind(this);
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onDispose = this.dispose;
        this.properties.onRender = this.render;
        this.label = _properties.label;
        this.context = _properties.context;
        this.onPropertyChange = _properties.onPropertyChange;
        this.key = _properties.key;
        this.onGetErrorMessage = _properties.onGetErrorMessage;
        this.panelTitle = _properties.panelTitle;
        this.limitByGroupNameOrID = _properties.limitByGroupNameOrID;
        this.limitByTermsetNameOrID = _properties.limitByTermsetNameOrID;
        this.hideTermStoreName = _properties.hideTermStoreName;
        this.isTermSetSelectable = _properties.isTermSetSelectable;
        this.disabledTermIds = _properties.disabledTermIds;
        this.termService = _properties.termService;
        this.spService = _properties.spService;
        this.numberOfZones = _properties.numberOfZones || 6;

        if (_properties.disabled === true) {
            this.disabled = _properties.disabled;
        }
        if (_properties.deferredValidationTime) {
            this.deferredValidationTime = _properties.deferredValidationTime;
        }
        if (typeof _properties.allowMultipleSelections !== 'undefined') {
            this.allowMultipleSelections = _properties.allowMultipleSelections;
        }
        if (typeof _properties.value !== "undefined" && _properties.value !== null) {
            this.propertyValue = _properties.value;
        }
        if (typeof _properties.excludeSystemGroup !== 'undefined') {
            this.excludeSystemGroup = _properties.excludeSystemGroup;
        }
    }

    /**
     * Renders the Multi Zone News Selector field content
     */
    private render(elem: HTMLElement, ctx?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        // Construct the JSX properties
        const element: React.ReactElement<IPropertyFieldMultiZoneSelectorHostProps> = React.createElement(PropertyFieldMultiZoneNewsSelectorHost, {
            label: this.label,
            targetProperty: this.targetProperty,
            panelTitle: this.panelTitle,
            allowMultipleSelections: this.allowMultipleSelections,
            numberOfZones: this.numberOfZones,
            value: this.propertyValue,
            excludeSystemGroup: this.excludeSystemGroup,
            limitByGroupNameOrID: this.limitByGroupNameOrID,
            limitByTermsetNameOrID: this.limitByTermsetNameOrID,
            hideTermStoreName: this.hideTermStoreName,
            isTermSetSelectable: this.isTermSetSelectable,
            disabledTermIds: this.disabledTermIds,
            context: this.context,
            onDispose: this.dispose,
            onRender: this.render,
            onChange: changeCallback,
            onPropertyChange: this.onPropertyChange,
            key: this.key,
            disabled: this.disabled,
            onGetErrorMessage: this.onGetErrorMessage,
            deferredValidationTime: this.deferredValidationTime,
            termService: this.termService,
            spService: this.spService
        });

        // Calls the REACT content generator
        ReactDom.render(element, elem);
    }

    /**
     * Disposes the current object
     */
    private dispose(elem: HTMLElement): void {
    }
}

/**
 * Helper method to create a SPList Picker on the PropertyPane.
 * @param targetProperty - Target property the SharePoint list picker is associated to.
 * @param properties - Strongly typed SPList Picker properties.
 */
export function PropertyFieldMultiZoneNewsSelector(targetProperty: string, properties: IPropertyFieldMultiZoneSelectorProps): IPropertyPaneField<IPropertyFieldMultiZoneSelectorPropsInternal> {
    // Calls the PropertyFieldTermPicker builder object
    // This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyPaneMultiZoneNewsSelectorBuilder(targetProperty, {
        ...properties,
        targetProperty: targetProperty,
        onRender: null,
        onDispose: null,
        termService: new SPTermStorePickerService(properties, properties.context),
        spService: new SPService(properties.context)
    });
}