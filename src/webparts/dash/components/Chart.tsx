import * as React from 'react';
import { IListItem } from '../../../services/SharePoint/IListItem';
import SharePointService from '../../../services/SharePoint/SharePointService';
import {
  Bar,
  Line,
  HorizontalBar,
  Pie,
  Doughnut,
} from 'react-chartjs-2';

export interface IChartProps {
  chartTitle: string;
}

export interface IChartState {
  items: IListItem[];
  loading: boolean;
  error: string | null;
}

export default class Chart extends React.Component<IChartProps, IChartState> {
  constructor(props: IChartProps) {
    super(props);

    // Bind methods
    this.getItems = this.getItems.bind(this);
    this.chartData = this.chartData.bind(this);

    // Set initial state
    this.state = {
      items: [],
      loading: false,
      error: null,
    };
  }

  public render(): JSX.Element {
    return (
      <div>
        <h1>{this.props.chartTitle}</h1>

        {this.state.error && <p>{this.state.error}</p>}

        <Bar data={this.chartData()} />

        <button onClick={this.getItems} disabled={this.state.loading}>
          {this.state.loading ? 'Loading...' : 'Refresh'}
        </button>
      </div>
    );
  }

  public getItems(): void {
    this.setState({ loading: true });

    SharePointService.getListItems('9fe2fbea-d7e9-4123-8f03-6bbea967b034').then(items => {
      this.setState({
        items: items.value,
        loading: false,
        error: null,
      });
    }).catch(error => {
      this.setState({
        error: 'Something went wrong!',
        loading: false,
      });
    });
  }

  public chartData(): object {
    const colors = [
      '#0078d4',
      '#bad80a',
      '#00b294',
      '#5c2d91',
      '#e3008c',
    ];

    // Chart data
    const data = {
      labels: [
        'Q1',
        'Q2',
        'Q3',
        'Q4',
      ],
      datasets: [],
    };

    // Add datasets
    this.state.items.map((item, i) => {
      // Create dataset
      const dataset = {
        label: item.Title,
        data: [
          item.EarningsQ1,
          item.EarningsQ2,
          item.EarningsQ3,
          item.EarningsQ4,
        ],
        backgroundColor: colors[i % colors.length],
        borderColor: colors[i % colors.length],
      };

      data.datasets.push(dataset);
    });

    return data;
  }
}
