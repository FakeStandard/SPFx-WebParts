import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./ValidatePropertyValuesWebPart.module.scss";
import * as strings from "ValidatePropertyValuesWebPartStrings";

// import
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IValidatePropertyValuesWebPartProps {
  description: string;
  listName: string;
}

export default class ValidatePropertyValuesWebPart extends BaseClientSideWebPart<IValidatePropertyValuesWebPartProps> {
  // add validation method (internal)
  private validateDescription(value: string): string {
    if (value === null || value.trim().length === 0)
      return "Provide a description";

    if (value.length > 10) {
      return "Description should not be longer than 10 characters";
    }
  }

  // add remote api validation method(external)
  private async validateListName(value: string): Promise<string> {
    if (value === null || value.length === 0) {
      return "Provide the list name";
    }

    try {
      let response = await this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        return "";
      } else if (response.status === 404) {
        return `List '${escape(value)}' doesn't exist in the current site`;
      } else {
        return `Error: ${response.statusText}. Please try again`;
      }
    } catch (error) {
      return error.message;
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.validatePropertyValues}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${
                styles.subTitle
              }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
                  // add validate
                  onGetErrorMessage: this.validateDescription.bind(this),
                }),
                PropertyPaneTextField("listName", {
                  label: strings.ListNameFieldLabel,
                  onGetErrorMessage: this.validateListName.bind(this),
                  // validatioin delay
                  deferredValidationTime: 500,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
