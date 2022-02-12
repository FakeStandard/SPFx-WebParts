import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "CustomPropertyPaneControlWebPartStrings";
import CustomPropertyPaneControl from "./components/CustomPropertyPaneControl";
import { ICustomPropertyPaneControlProps } from "./components/ICustomPropertyPaneControlProps";

import { PropertyPaneAsyncDropdown } from "../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown";
import { IDropdownOption } from "office-ui-fabric-react/lib/components/Dropdown";
import { update, get } from "@microsoft/sp-lodash-subset";

export interface ICustomPropertyPaneControlWebPartProps {
  listName: string;
  item: string;
}

export interface IListInfo {
  Id: string;
  Title: string;
}

export default class CustomPropertyPaneControlWebPart extends BaseClientSideWebPart<ICustomPropertyPaneControlWebPartProps> {
  public render(): void {
    // const element: React.ReactElement<ICustomPropertyPaneControlProps> =
    //   React.createElement(CustomPropertyPaneControl, {
    //     description: this.properties.description,
    //   });

    // ReactDom.render(element, this.domElement);

    const element: React.ReactElement<ICustomPropertyPaneControlWebPartProps> =
      React.createElement(CustomPropertyPaneControl, {
        listName: this.properties.listName,
        item: this.properties.item,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  private itemsDropDown: PropertyPaneAsyncDropdown;

  private onListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => {
      return newValue;
    });
    // reset selected item
    this.properties.item = undefined;
    // store new value in web part properties
    update(this.properties, "item", (): any => {
      return this.properties.item;
    });
    // refresh web part
    this.render();
    // reset selected values in item dropdown
    this.itemsDropDown.properties.selectedKey = this.properties.item;
    // allow to load items
    this.itemsDropDown.properties.disabled = false;
    // load items and re-render items dropdown
    this.itemsDropDown.render();
  }

  private loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(
      (
        resolve: (options: IDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        setTimeout(() => {
          resolve([
            {
              key: "sharedDocuments",
              text: "Shared Documents",
            },
            {
              key: "myDocuments",
              text: "My Documents",
            },
          ]);
        }, 2000);
      }
    );
  }

  private loadItems(): Promise<IDropdownOption[]> {
    // if (!this.properties.listName) {
    //   // resolve to empty options since no list has been selected
    //   return null;
    // }

    const wp: CustomPropertyPaneControlWebPart = this;

    return new Promise<IDropdownOption[]>(
      (
        resolve: (options: IDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        setTimeout(() => {
          const items = {
            sharedDocuments: [
              {
                key: "spfx_presentation.pptx",
                text: "SPFx for the masses",
              },
              {
                key: "hello-world.spapp",
                text: "hello-world.spapp",
              },
            ],
            myDocuments: [
              {
                key: "isaiah_cv.docx",
                text: "Isaiah CV",
              },
              {
                key: "isaiah_expenses.xlsx",
                text: "Isaiah Expenses",
              },
            ],
          };
          resolve(items[wp.properties.listName]);
        }, 2000);
      }
    );
  }

  private onListItemChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => {
      return newValue;
    });
    // refresh web part
    this.render();
  }

  // private onListChange(propertyPath: string, newValue: any): void {
  //   const oldValue: any = get(this.properties, propertyPath);
  //   // store new value in web part properties
  //   update(this.properties, propertyPath, (): any => {
  //     return newValue;
  //   });
  //   // refresh web part
  //   this.render();
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // reference to item dropdown needed later after selecting a list
    this.itemsDropDown = new PropertyPaneAsyncDropdown("item", {
      label: strings.ItemFieldLabel,
      loadOptions: this.loadItems.bind(this),
      onPropertyChange: this.onListItemChange.bind(this),
      selectedKey: this.properties.item,
      // should be disabled if no list has been selected
      disabled: !this.properties.listName,
    });
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
                new PropertyPaneAsyncDropdown("listName", {
                  label: strings.ListFieldLabel,
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName,
                }),
                this.itemsDropDown,
              ],
            },
          ],
        },
      ],
    };
  }
}
function resolve(): any {
  throw new Error("Function not implemented.");
}

