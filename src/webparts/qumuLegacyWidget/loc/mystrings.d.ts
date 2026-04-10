declare interface IQumuLegacyWidgetWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  HostFieldLabel: string;
  GuidFieldLabel: string;
  TypeFieldLabel: string;
  LoadFailedPrefix: string;
  UnknownError: string;
}

declare module 'QumuLegacyWidgetWebPartStrings' {
  const strings: IQumuLegacyWidgetWebPartStrings;
  export = strings;
}
