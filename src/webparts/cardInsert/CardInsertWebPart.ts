import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import * as strings from 'CardInsertWebPartStrings';
import CardInsert from './components/CardInsert';
import {IPropsInsert} from '../../classes/IProps'

export interface ICardInsertWebPartProps {
  description: string;
  context: any;
  siteurl: string;
}

export default class CardInsertWebPart extends BaseClientSideWebPart<ICardInsertWebPartProps> {
  
  
  public render(): void {
    const element: React.ReactElement<IPropsInsert > = React.createElement(
      CardInsert,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.web.absoluteUrl
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
                  label: 'Header'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
