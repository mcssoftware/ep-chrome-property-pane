import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import { IPropertyFieldEpChromePropsInternal, IPropertyFieldEpChromeProps, IPropertyFieldEpChromeData } from './IPropertyFieldEpChrome';
import PropertPaneEpChromeHost from "./component/PropertyPaneEpChromeHost";

class PropertyFieldEpChromeBuilder implements IPropertyPaneField<IPropertyFieldEpChromePropsInternal> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyFieldEpChromePropsInternal;

    private _onChangeCallback: (targetProperty?: string, newValue?: any) => void;

    public constructor(_targetProperty: string, _properties: IPropertyFieldEpChromePropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
    }

    private _render(elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        const props: IPropertyFieldEpChromeProps = <IPropertyFieldEpChromeProps>this.properties;
        const element = React.createElement(PropertPaneEpChromeHost, {
            ...props,
            targetProperty: this.targetProperty,
            onChange: this._onChanged.bind(this)
        });
        ReactDOM.render(element, elem);
        if (changeCallback) {
            this._onChangeCallback = changeCallback;
        }
    }

    private _dispose(elem: HTMLElement) {
        ReactDOM.unmountComponentAtNode(elem);
    }

    private _onChanged(targetProperty?: string, value?: IPropertyFieldEpChromeData): void {
        if (this._onChangeCallback) {
            this._onChangeCallback(targetProperty, value);
        }
    }
}

export function PropertyFieldEpChrome(targetProperty: string, properties: IPropertyFieldEpChromeProps): IPropertyPaneField<IPropertyFieldEpChromePropsInternal> {
    return new PropertyFieldEpChromeBuilder(targetProperty, {
        ...properties,
        targetProperty: targetProperty,
        onRender: null,
        onDispose: null
    });
}
