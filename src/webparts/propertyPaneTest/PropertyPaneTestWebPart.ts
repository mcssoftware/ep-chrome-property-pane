import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PropertyPaneTestWebPartStrings';
import PropertyPaneTest from './components/PropertyPaneTest';
import { IPropertyPaneTestProps } from './components/IPropertyPaneTestProps';
import { IPropertyFieldEpChromeData, PropertyFieldEpChrome } from '../../propertyField/epchrome';
import { PropertyFieldNewsSelector, IPropertyFieldNewsSelectorData } from '../../propertyField/newsSelector';
import { get, update } from '@microsoft/sp-lodash-subset';

export interface IPropertyPaneTestWebPartProps {
  description: string;
  test: IPropertyFieldEpChromeData;
  newsSelector: IPropertyFieldNewsSelectorData;
}

export default class PropertyPaneTestWebPart extends BaseClientSideWebPart<IPropertyPaneTestWebPartProps> {

  /**
   *
   */
  constructor() {
    super();
    this._validation = this._validation.bind(this);
    this.onPropertyChanged = this.onPropertyChanged.bind(this);
  }

  public render(): void {
    const element: React.ReactElement<IPropertyPaneTestProps> = React.createElement(
      PropertyPaneTest,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "News Settings"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyFieldEpChrome("test", {
                  key: "test",
                  value: this.properties.test,
                  label: "Test Settings",
                  onPropertyChange: this.onPropertyChanged,
                }),
                PropertyFieldNewsSelector("newsSelector", {
                  key: "newsSelector",
                  context: this.context,
                  allowMultipleSelections: false,
                  disabled: false,
                  hideTermStoreName: true,
                  label: "News Selector",
                  onGetErrorMessage: this._validation,
                  panelTitle: "News Selector Panel",
                  limitByGroupNameOrID: "ElevatePoint",
                  limitByTermsetNameOrID: "Department",
                  properties: {},
                  onPropertyChange: this.onPropertyChanged
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private onPropertyChanged(propertyPath: string, _oldValue: any, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // refresh web part
    this.render();
  }

  private _validation(value: IPropertyFieldNewsSelectorData): string {
    return "";
  }
}
