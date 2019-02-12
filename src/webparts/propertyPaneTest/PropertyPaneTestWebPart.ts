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

export interface IPropertyPaneTestWebPartProps {
  description: string;
  test: IPropertyFieldEpChromeData;
}

export default class PropertyPaneTestWebPart extends BaseClientSideWebPart<IPropertyPaneTestWebPartProps> {

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
              groupName: "Chrome Settings",
              groupFields: [
                PropertyFieldEpChrome('test', {
                  key: "test",
                  value: this.properties.test
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
