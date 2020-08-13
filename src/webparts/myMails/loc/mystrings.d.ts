declare interface IMyMailsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ApplicationIDFieldLabel: string;
  RedirectUriFieldLabel: string;
  TenantUrlFieldLabel: string;
}

declare module 'MyMailsWebPartStrings' {
  const strings: IMyMailsWebPartStrings;
  export = strings;
}
