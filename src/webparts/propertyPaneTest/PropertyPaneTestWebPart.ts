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
import { IPropertyFieldCalendarData, PropertyFieldCalendar } from '../../propertyField/calendar';
import { ListPickerOrderByType } from '../../services/ISPService';
import { IPropertyPaneMultiZoneSelectorData, PropertyFieldMultiZoneNewsSelector } from '../../propertyField/multiZoneSelector';

export interface IPropertyPaneTestWebPartProps {
  description: string;
  test: IPropertyFieldEpChromeData;
  innerContent: string;
  newsSelector: IPropertyFieldNewsSelectorData;
  calendarSelector: IPropertyFieldCalendarData;
  multizoneNewsSelector: IPropertyPaneMultiZoneSelectorData;
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
                PropertyFieldMultiZoneNewsSelector("multizoneNewsSelector", {
                  key: "multizoneNewsSelector",
                  context: this.context,
                  allowMultipleSelections: false,
                  disabled: false,
                  hideTermStoreName: true,
                  label: "Zone News Selector",
                  onGetErrorMessage: null, //this._validation,
                  panelTitle: "News Selector Panel",
                  limitByGroupNameOrID: "ElevatePoint",
                  limitByTermsetNameOrID: "News Channel",
                  onPropertyChange: this.onPropertyChanged,
                  value: this.properties.multizoneNewsSelector,
                  numberOfZones: 6
                })
              ]
            }
          ]
        },
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
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "News Settings"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
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
                  limitByTermsetNameOrID: "News Channel",
                  onPropertyChange: this.onPropertyChanged
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "News Settings"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyFieldCalendar("calendarSelector", {
                  key: "calendarSelector",
                  context: this.context,
                  includeHiddenList: false,
                  label: "Calendar",
                  onGetErrorMessage: null,
                  listBaseTemplate: 106,
                  value: this.properties.calendarSelector,
                  onPropertyChange: this.onPropertyChanged,
                  listOrderBy: ListPickerOrderByType.Title
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
