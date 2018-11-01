declare interface IDashWebPartStrings {
  AddColor: string;
  ChartBar: string;
  ChartBarHorizontal: string;
  ChartData: string;
  ChartDonut: string;
  ChartLine: string;
  ChartPie: string;
  ChartSettings: string;
  ChartStyle: string;
  ChartTitle: string;
  ChartType: string;
  Colors: string;
  DeleteColor: string;
  Error: string;
  Intro: string;
  List: string;
  Loading: string;
  LoadingChartData: string;
  PropertyPaneDescription: string;
  Refresh: string;
  SelectedFields: string;
}

declare module 'DashWebPartStrings' {
  const strings: IDashWebPartStrings;
  export = strings;
}
