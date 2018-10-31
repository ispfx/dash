import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DashWebPartStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IDashWebPartProps {
  description: string;
}

export default class DashWebPart extends BaseClientSideWebPart<IDashWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashProps > = React.createElement(
      Dash,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      if (Environment.type == EnvironmentType.Local) {
        // return MockData;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getbytitle('Produce Revenue')/items?$select=Title,January,February&$filter=February gt 20`, SPHttpClient.configurations.v1).then(response => {
        response.json().then((json: any) => {
          console.log(json);
        });
      });
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
