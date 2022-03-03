import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ReactCrudWebPartStrings";
import ReactCrud from "./components/ReactCrud";
import { IReactCrudProps } from "./components/IReactCrudProps";

export interface IReactCrudWebPartProps {
  description: string;
}

export default class ReactCrudWebPart extends BaseClientSideWebPart<IReactCrudWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IReactCrudProps> = React.createElement(
      ReactCrud,
      {
        description: this.properties.description,
        context: this.context,
        webURL: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
