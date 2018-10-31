import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import * as strings from 'DashWebPartStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import SharePointService from '../../services/SharePoint/SharePointService';

export interface IDashWebPartProps {
  listId: string;
  selectedFields: string;
  chartType: string;
  chartTitle: string;
  color1: string;
  color2: string;
  color3: string;
}

export default class DashWebPart extends BaseClientSideWebPart<IDashWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashProps > = React.createElement(
      Dash,
      {
        listId: this.properties.listId,
        selectedFields: this.properties.selectedFields.split(','),
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
            description: 'Dash Settings'
          },
          groups: [
            {
              groupName: 'Chart Data',
              groupFields: [
                PropertyPaneTextField('listId', {
                  label: 'List'
                }),
                PropertyPaneTextField('selectedFields', {
                  label: 'Selected Fields'
                }),
              ]
            },
            {
              groupName: 'Chart Settings',
              groupFields: [
                PropertyPaneDropdown('chartType', {
                  label: 'Chart Type',
                  options: [
                    { key: 'Bar', text: 'Bar' },
                    { key: 'HorizontalBar', text: 'HorizontalBar' },
                    { key: 'Line', text: 'Line' },
                    { key: 'Pie', text: 'Pie' },
                    { key: 'Doughnut', text: 'Doughnut' },
                  ],
                }),
                PropertyPaneTextField('chartTitle', {
                  label: 'Chart Title'
                }),
              ],
            },
            {
              groupName: 'Chart Style',
              groupFields: [
                PropertyFieldColorPicker('color1', {
                  label: 'Color 1',
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
                  label: 'Color 2',
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
                  label: 'Color 3',
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
}
