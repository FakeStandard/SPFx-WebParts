import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PeoplepickerSampleWebPartStrings';
import PeoplepickerSample from './components/PeoplepickerSample';
import { IPeoplepickerSampleProps } from './components/IPeoplepickerSampleProps';

export interface IPeoplepickerSampleWebPartProps {
  description: string;
}

export default class PeoplepickerSampleWebPart extends BaseClientSideWebPart<IPeoplepickerSampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPeoplepickerSampleProps> = React.createElement(
      PeoplepickerSample,
      {
        description: this.properties.description,
        context: this.context
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
