import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TemperatureConversionWebPartStrings';
import TemperatureConversion from './components/TemperatureConversion';
import { ITemperatureConversionProps } from './components/ITemperatureConversionProps';

export interface ITemperatureConversionWebPartProps {
  description: string;
}

export default class TemperatureConversionWebPart extends BaseClientSideWebPart<ITemperatureConversionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITemperatureConversionProps> = React.createElement(
      TemperatureConversion,
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
