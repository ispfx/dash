import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

import { PropertyPaneColorPalette } from '../../controls/PropertyPaneColorPalette/PropertyPaneColorPalette';
import * as strings from 'DashWebPartStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import SharePointService from '../../services/SharePoint/SharePointService';
import { updateA } from 'office-ui-fabric-react/lib/utilities/color';

export interface IDashWebPartProps {
  listId: string;
  selectedFields: string[];
  chartType: string;
  chartTitle: string;
  colors: string[];
}

export default class DashWebPart extends BaseClientSideWebPart<IDashWebPartProps> {
  // List options state
  private listOptions: IPropertyPaneDropdownOption[];
  private listOptionsLoading: boolean = false;

  // Field options state
  private fieldOptions: IPropertyPaneDropdownOption[];
  private fieldOptionsLoading: boolean = false;

  public render(): void {
    const element: React.ReactElement<IDashProps > = React.createElement(
      Dash,
      {
        listId: this.properties.listId,
        selectedFields: this.properties.selectedFields,
        chartType: this.properties.chartType,
        chartTitle: this.properties.chartTitle,
        colors: this.properties.colors,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      SharePointService.setup(this.context, Environment.type);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.ChartData,
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: strings.List,
                  options: this.listOptions,
                  disabled: this.listOptionsLoading,
                }),
                PropertyFieldMultiSelect('selectedFields', {
                  key: 'selectedFields',
                  label: strings.SelectedFields,
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading,
                  selectedKeys: this.properties.selectedFields,
                })
              ]
            },
            {
              groupName: strings.ChartSettings,
              groupFields: [
                PropertyPaneDropdown('chartType', {
                  label: strings.ChartType,
                  options: [
                    { key: 'Bar', text: strings.ChartBar },
                    { key: 'HorizontalBar', text: strings.ChartBarHorizontal },
                    { key: 'Line', text: strings.ChartLine },
                    { key: 'Pie', text: strings.ChartPie },
                    { key: 'Doughnut', text: strings.ChartDonut },
                  ],
                }),
                PropertyPaneTextField('chartTitle', {
                  label: strings.ChartTitle
                }),
              ],
            },
            {
              groupName: strings.ChartStyle,
              groupFields: [
                new PropertyPaneColorPalette('colors', {
                  label: strings.Colors,
                  colors: this.properties.colors,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  key: 'colors_palette',
                }),
              ],
            }
          ]
        }
      ]
    };
  }

  private getLists(): Promise<IPropertyPaneDropdownOption[]> {
    this.listOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getLists().then(lists => {
      this.listOptionsLoading = false;
      this.context.propertyPane.refresh();

      return lists.value.map(list => {
        return {
          key: list.Id,
          text: list.Title,
        };
      });
    });
  }

  public getFields(): Promise<IPropertyPaneDropdownOption[]> {
    // No list selected
    if (!this.properties.listId) return Promise.resolve();

    this.fieldOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getListFields(this.properties.listId).then(fields => {
      this.fieldOptionsLoading = false;
      this.context.propertyPane.refresh();

      return fields.value.map(field => {
        return {
          key: field.InternalName,
          text: `${field.Title} (${field.TypeAsString})`,
        };
      });
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.getLists().then(listOptions => {
      this.listOptions = listOptions;
      this.context.propertyPane.refresh();
    }).then(() => {
      this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && newValue) {
      this.properties.selectedFields = [];

      this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    }

    else if (propertyPath === 'colors' && newValue) {
      this.properties.colors = newValue;
      this.context.propertyPane.refresh();
      this.render();
    }
  }
}
