import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import { IPropertyFieldKeyEventsPropsInternal, IPropertyFieldKeyEventsProps } from './IPropertyFieldKeyEvents';
import PropertyFieldKeyEventsHost from "./component/PropertyPaneKeyEventsHost";
import SPService from '../../services/SPService';

class PropertyFieldKeyEventsBuilder implements IPropertyPaneField<IPropertyFieldKeyEventsPropsInternal> {

    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyFieldKeyEventsPropsInternal;
    private deferredValidationTime: number = 200;
    private disabled: boolean = false;

    public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void { }
    // private _onChangeCallback: (targetProperty?: string, newValue?: any) => void;

    public constructor(_targetProperty: string, _properties: IPropertyFieldKeyEventsPropsInternal) {
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
        const element = React.createElement(PropertyFieldKeyEventsHost, {
            ...this.properties,
            // key: this.properties.key,
            // label: this.properties.label,
            // value: this.properties.value,
            disabled: this.disabled,
            // context: this.properties.context,
            onChange: changeCallback,
            // onGetErrorMessage: this.properties.onGetErrorMessage,
            onPropertyChange: this.onPropertyChange,
            deferredValidationTime: this.deferredValidationTime,
        });
        ReactDOM.render(element, elem);
        // if (changeCallback) {
        //     this._onChangeCallback = changeCallback;
        // }
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

export function PropertyFieldKeyEvents(targetProperty: string, properties: IPropertyFieldKeyEventsProps): IPropertyPaneField<IPropertyFieldKeyEventsPropsInternal> {
    return new PropertyFieldKeyEventsBuilder(targetProperty, {
        ...properties,
        targetProperty: targetProperty,
        onRender: null,
        onDispose: null,
        spService: new SPService(properties.context)
    });
}