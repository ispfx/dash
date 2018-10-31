import * as React from 'react';
import styles from './Dash.module.scss';
import { IDashProps } from './IDashProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Chart from './Chart';

export default class Dash extends React.Component<IDashProps, {}> {
  public render(): React.ReactElement<IDashProps> {
    return (
      <div className={ styles.dash }>
        <Chart
          listId={this.props.listId}
          selectedFields={this.props.selectedFields}
          chartType={this.props.chartType}
          chartTitle={this.props.chartTitle}
          colors={this.props.colors} />
      </div>
    );
  }
}
