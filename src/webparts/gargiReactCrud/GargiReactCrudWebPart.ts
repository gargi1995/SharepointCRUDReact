import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GargiReactCrudWebPartStrings';
import GargiReactCrud from './components/GargiReactCrud';
import { IGargiReactCrudProps } from './components/IGargiReactCrudProps';

export interface IGargiReactCrudWebPartProps {
  description: string;  
}

export default class GargiReactCrudWebPart extends BaseClientSideWebPart <IGargiReactCrudWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGargiReactCrudProps> = React.createElement(
      GargiReactCrud,
      {
        description: this.properties.description,
        context:this.context
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
                  label: strings.DescriptionFieldLabel})
              ]
            }
          ]
        }
      ]
    };
  }
}
