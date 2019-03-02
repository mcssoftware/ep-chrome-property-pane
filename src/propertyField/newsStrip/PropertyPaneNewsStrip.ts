import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IPropertyFieldNewsStripProps, IPropertyFieldNewsStripPropsInternal } from './IPropertyPaneNewsStrip';
import { PropertyPaneFieldType, IPropertyPaneField } from '@microsoft/sp-webpart-base';
import PropertyFieldNewsStripHost from "./component/PropertyFieldNewsStripHost";

class PropertyFieldNewsStripBuilder implements IPropertyPaneField<IPropertyFieldNewsStripPropsInternal> {

    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyFieldNewsStripPropsInternal;
    private deferredValidationTime: number = 200;
    private disabled: boolean = false;

    public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void { }
    // private _onChangeCallback: (targetProperty?: string, newValue?: any) => void;

    public constructor(_targetProperty: string, _properties: IPropertyFieldNewsStripPropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
        this.onPropertyChange = _properties.onPropertyChange;
        if (_properties.disabled === true) {
            this.disabled = _properties.disabled;
        }
        if (_properties.deferredValidationTime) {
            this.deferredValidationTime = _properties.deferredValidationTime;
        }
    }

    private _render(elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        const element = React.createElement(PropertyFieldNewsStripHost, {
            ...this.properties,
            onChange: changeCallback,
            disabled: this.disabled,
            onPropertyChange: this.onPropertyChange,
            deferredValidationTime: this.deferredValidationTime,
        });
        ReactDOM.render(element, elem);
    }

    private _dispose(elem: HTMLElement) {
        ReactDOM.unmountComponentAtNode(elem);
    }

    // private _onChanged(targetProperty?: string, value?: IPropertyPaneKeyEventsData): void {
    //     if (this._onChangeCallback) {
    //         this._onChangeCallback(targetProperty, value);
    //     }
    // }
}

export function PropertyFieldNewsStrip(targetProperty: string, properties: IPropertyFieldNewsStripProps): IPropertyPaneField<IPropertyFieldNewsStripPropsInternal> {
    return new PropertyFieldNewsStripBuilder(targetProperty, {
        ...properties,
        targetProperty: targetProperty,
        onRender: null,
        onDispose: null,
    });
}