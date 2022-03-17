import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'IconSampleWebPartStrings';
import IconSample from './components/IconSample';
import { IIconSampleProps } from './components/IIconSampleProps';

export interface IIconSampleWebPartProps {
  description: string;
}

export default class IconSampleWebPart extends BaseClientSideWebPart<IIconSampleWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      (document.querySelector("#workbenchPageContent") as HTMLElement).style.maxWidth =
        "none";
      (document.querySelector(".CanvasZone") as HTMLElement).style.maxWidth =
        "none";
    });
  }

  public render(): void {
    const element: React.ReactElement<IIconSampleProps> = React.createElement(
      IconSample,
      {
        // description: this.properties.description
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
