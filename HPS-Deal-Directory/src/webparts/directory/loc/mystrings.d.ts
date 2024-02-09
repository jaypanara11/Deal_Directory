declare interface IDirectoryWebPartStrings {
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
  ClearTextSearchPropsLabel: string;
  ClearTextSearchPropsDesc: string;
  PagingLabel: string;
}

declare module 'DirectoryWebPartStrings' {
  const strings: IDirectoryWebPartStrings;
  export = strings;
}
