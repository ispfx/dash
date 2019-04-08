import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IDashProps {
  listId: string;
  selectedFields: string[];
  chartType: string;
  chartTitle: string;
  colors: string[];
  theme: IReadonlyTheme | undefined;
}
