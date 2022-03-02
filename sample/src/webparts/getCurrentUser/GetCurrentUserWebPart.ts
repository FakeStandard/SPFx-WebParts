import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./GetCurrentUserWebPart.module.scss";
import * as strings from "GetCurrentUserWebPartStrings";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { ISiteUser, ISiteUserInfo } from "@pnp/sp/site-users/types";
import { ThemeProvider } from "office-ui-fabric-react/lib/Foundation";
import { ISiteGroups } from "@pnp/sp/site-groups/types";

export interface IGetCurrentUserWebPartProps {
  description: string;
}

export default class GetCurrentUserWebPart extends BaseClientSideWebPart<IGetCurrentUserWebPartProps> {
  //Get Current User Display Name
  private async getSPData(): Promise<void> {
    await sp.web.currentUser.get().then((r: ISiteUserInfo) => {
      this.renderData(r["Title"]);
    });
  }

  private async getSPGroup(): Promise<void> {
    await sp.web.currentUser.groups.get().then((r: any) => {
      let grpNames: string = "";
      r.forEach((grp: ISiteGroups) => {
        grpNames += "<li>" + grp["Title"] + "</li>";
      });
      grpNames = "<ul>" + grpNames + "</ul>";
      this.renderGroupData(grpNames);
    });
  }

  private renderData(strResponse: string): void {
    const htmlElement = this.domElement.querySelector("#pnpinfo");
    htmlElement.innerHTML = strResponse;
  }

  private renderGroupData(strResponse: string): void {
    const htmlElement = this.domElement.querySelector("#pnpGroup");
    htmlElement.innerHTML = strResponse;
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `  
    <div class="${styles.getCurrentUser}">  
      <div class="${styles.container}">  
        <div class="${styles.row}">  
          <div class="${styles.column}">  
            <div id="pnpinfo"></div>  
            <div id="pnpGroup"></div>  
          </div>  
        </div>  
      </div>  
    </div>`;
    this.getSPData();
    this.getSPGroup();
    this.getCurrentUser();
  }

  private getCurrentUser(): void {
    let name = this.context.pageContext.user.displayName;
    let mail = this.context.pageContext.user.email;

    console.log(name);

    sp.site.rootWeb.ensureUser(mail).then((result) => {
      console.log(result.data.Id);
    });
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
