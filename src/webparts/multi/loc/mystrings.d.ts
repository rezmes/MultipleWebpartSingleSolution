declare interface IMultiWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MultiWebPartStrings' {
  const strings: IMultiWebPartStrings;
  export = strings;
}
