declare interface IConfigurePropertyPaneWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  PropertyPaneCheckbox: boolean;
  PropertyPaneDropdown: string;
  PropertyPaneToggle: boolean;
}

declare module "ConfigurePropertyPaneWebPartStrings" {
  const strings: IConfigurePropertyPaneWebPartStrings;
  export = strings;
}
