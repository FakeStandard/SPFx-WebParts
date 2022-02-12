declare interface IValidatePropertyValuesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListNameFieldLabel: string;
}

declare module "ValidatePropertyValuesWebPartStrings" {
  const strings: IValidatePropertyValuesWebPartStrings;
  export = strings;
}
