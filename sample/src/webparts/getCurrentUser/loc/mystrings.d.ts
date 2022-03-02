declare interface IGetCurrentUserWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'GetCurrentUserWebPartStrings' {
  const strings: IGetCurrentUserWebPartStrings;
  export = strings;
}
