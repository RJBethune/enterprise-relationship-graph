declare interface IGraphWebPartStrings {
  PropertyPaneDescription: string;
  DataGroupName: string;
  SyncGroupName: string;
  AdvancedGroupName: string;
  ProjectIdFieldLabel: string;
  PollSecondsFieldLabel: string;
  AutosaveFieldLabel: string;
  AssetBaseUrlFieldLabel: string;
}

declare module 'GraphWebPartStrings' {
  const strings: IGraphWebPartStrings;
  export = strings;
}
