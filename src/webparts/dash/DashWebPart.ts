import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
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

import * as strings from 'DashWebPartStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import SharePointService from '../../services/SharePoint/SharePointService';

export interface IDashWebPartProps {
  listId: string;
  selectedFields: string[];
  chartType: string;
  chartTitle: string;
  color1: string;
  color2: string;
  color3: string;
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
        colors: [
          this.properties.color1,
          this.properties.color2,
          this.properties.color3,
        ],
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
                PropertyFieldColorPicker('color1', {
                  label: strings.Colors,
                  selectedColor: this.properties.color1,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker1'
                }),
                PropertyFieldColorPicker('color2', {
                  label: strings.Colors,
                  selectedColor: this.properties.color2,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker2'
                }),
                PropertyFieldColorPicker('color3', {
                  label: strings.Colors,
                  selectedColor: this.properties.color3,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker3'
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
  }
}
