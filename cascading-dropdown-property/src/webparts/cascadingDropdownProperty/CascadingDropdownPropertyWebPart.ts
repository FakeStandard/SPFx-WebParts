import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "CascadingDropdownPropertyWebPartStrings";
import CascadingDropdownProperty from "./components/CascadingDropdownProperty";
import { ICascadingDropdownPropertyProps } from "./components/ICascadingDropdownPropertyProps";

export interface ICascadingDropdownPropertyWebPartProps {
  description: string;
  listName: string;
  itemName: string;
}

export default class CascadingDropdownPropertyWebPart extends BaseClientSideWebPart<ICascadingDropdownPropertyWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICascadingDropdownPropertyProps> =
      React.createElement(CascadingDropdownProperty, {
        description: this.properties.description,
        listName: this.properties.listName,
        itemName: this.properties.itemName,
      });

    ReactDom.render(element, this.domElement);
  }

  private items: IPropertyPaneDropdownOption[];
  private itemsDropdownDisabled: boolean = true;

  private loadItems(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return null;
    }

    const wp: CascadingDropdownPropertyWebPart = this;

    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
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

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;
    this.itemsDropdownDisabled = !this.properties.listName || !this.items;

    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "options"
    );

    this.loadLists()
      .then(
        (
          listOptions: IPropertyPaneDropdownOption[]
        ): Promise<IPropertyPaneDropdownOption[]> => {
          this.lists = listOptions;
          this.listsDropdownDisabled = false;
          this.context.propertyPane.refresh();
          return this.loadItems();
        }
      )
      .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
        this.items = itemOptions;
        this.itemsDropdownDisabled = !this.properties.listName;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "listName" && newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousItem: string = this.properties.itemName;
      // reset selected item
      this.properties.itemName = undefined;
      // push new item value
      this.onPropertyPaneFieldChanged(
        "itemName",
        previousItem,
        this.properties.itemName
      );
      // disable item selector until new items are loaded
      this.itemsDropdownDisabled = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "items"
      );

      this.loadItems().then(
        (itemOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.items = itemOptions;
          // enable item selector
          this.itemsDropdownDisabled = false;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        }
      );
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        setTimeout((): void => {
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

  // protected onPropertyPaneConfigurationStart(): void {
  //   this.listsDropdownDisabled = !this.lists;

  //   if (this.lists) {
  //     return;
  //   }

  //   this.context.statusRenderer.displayLoadingIndicator(
  //     this.domElement,
  //     "lists"
  //   );

  //   this.loadLists().then(
  //     (listOptions: IPropertyPaneDropdownOption[]): void => {
  //       this.lists = listOptions;
  //       this.listsDropdownDisabled = false;
  //       this.context.propertyPane.refresh();
  //       this.context.statusRenderer.clearLoadingIndicator(this.domElement);
  //       this.render();
  //     }
  //   );
  // }

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
                PropertyPaneDropdown("listName", {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled,
                }),
                PropertyPaneDropdown("itemName", {
                  label: strings.ItemNameFieldLabel,
                  options: this.items,
                  disabled: this.itemsDropdownDisabled,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
