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

export interface IPropertyPaneTestWebPartProps {
  description: string;
}

export default class PropertyPaneTestWebPart extends BaseClientSideWebPart<IPropertyPaneTestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPropertyPaneTestProps > = React.createElement(
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
