declare interface IHelloWorldPracticeWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWorldPracticeWebPartStrings' {
  const strings: IHelloWorldPracticeWebPartStrings;
  export = strings;
}
