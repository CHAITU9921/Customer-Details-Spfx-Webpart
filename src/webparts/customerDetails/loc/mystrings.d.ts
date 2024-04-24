declare interface ICustomerDetailsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'CustomerDetailsWebPartStrings' {
  const strings: ICustomerDetailsWebPartStrings;
  export = strings;
}
