import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';

import * as strings from 'CardDisplayWebPartStrings';
import CardDisplay from './components/CardDisplay';
import { IProps } from '../../classes/IProps';

export interface ICardDisplayWebPartProps {
  listname: string;
  context: any;
  resturl: string;
  itemcount: number;
  assetfolderurl: string;
}

export default class CardDisplayWebPart extends BaseClientSideWebPart<ICardDisplayWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProps > = React.createElement(
      CardDisplay,
      {
        listname: this.properties.listname,
        context: this.context,
        resturl: this.properties.resturl,
        itemcount: this.properties.itemcount,
        assetfolderurl: this.properties.assetfolderurl
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
                PropertyPaneTextField('listname', {
                  label: "List Name",                  
                }),
                PropertyPaneTextField('resturl', {
                  label: "REST Url",                  
                }),
                PropertyPaneTextField('assetfolderurl', {
                  label: "Asset Folder Url",                  
                }),
                PropertyPaneSlider('itemcount', {
                  label: 'Number of Items',min:1,max:500
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
