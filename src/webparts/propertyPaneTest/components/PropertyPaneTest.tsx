import * as React from 'react';
import styles from './PropertyPaneTest.module.scss';
import { IPropertyPaneTestProps } from './IPropertyPaneTestProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PropertyPaneTest extends React.Component<IPropertyPaneTestProps, {}> {
  public render(): React.ReactElement<IPropertyPaneTestProps> {
    return (
      <div className={ styles.propertyPaneTest }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Webpart Properties</span>
              {JSON.stringify(this.props.properties)}
              https://cwsoft.sharepoint.com/sites/admc/Assets/Forms/AllItems.aspx
            </div>
          </div>
        </div>
      </div>
    );
  }
}
