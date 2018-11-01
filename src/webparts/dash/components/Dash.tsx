import * as React from 'react';
import { IDashProps } from './IDashProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Chart from './Chart';
import {
  MessageBar,
  MessageBarType,
} from 'office-ui-fabric-react/lib/MessageBar';
import * as strings from 'DashWebPartStrings';

export default class Dash extends React.Component<IDashProps, {}> {
  public render(): React.ReactElement<IDashProps> {
    return (
      <div>
        {this.props.listId ?
          <Chart
            listId={this.props.listId}
            selectedFields={this.props.selectedFields}
            chartType={this.props.chartType}
            chartTitle={this.props.chartTitle}
            colors={this.props.colors} /> :

          <MessageBar>{strings.Intro}</MessageBar>
        }
      </div>
    );
  }
}
