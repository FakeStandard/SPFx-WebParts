import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./DefineListModelWebPart.module.scss";
import * as strings from "DefineListModelWebPartStrings";

// import something
import MockHttpClient from "./MockHttpClient";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
//

export interface IDefineListModelWebPartProps {
  description: string;
}

// define the following two list interfaces
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class DefineListModelWebPart extends BaseClientSideWebPart<IDefineListModelWebPartProps> {
  // add
  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    } else if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  // add
  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    const listContainer: Element =
      this.domElement.querySelector("#spListContainer");
    listContainer.innerHTML = html;
  }

  // add
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // add
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get().then((data: ISPList[]) => {
      var listData: ISPLists = { value: data };
      return listData;
    }) as Promise<ISPLists>;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.defineListModel}">
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
              
              <p class="${styles.description}">Loading from ${escape(
      this.context.pageContext.web.title
    )}</p>

              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
          <div id="spListContainer" />
          </div>
      </div>`;

    this._renderListAsync();
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
