import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import { fmt } from '../types';
import styles from './SharePointSmartCharts.module.scss';

interface IChartErrorBoundaryProps {
  /** Changing this resets the boundary so a fixed configuration can re-render */
  resetKey: string;
}

interface IChartErrorBoundaryState {
  error: Error | null;
}

// Chart.js can throw during mount/update (after the renderer's own try/catch has
// returned) — without a boundary that unmounts the whole web part with an
// opaque SPFx error instead of a message the author can act on.
export default class ChartErrorBoundary extends React.Component<IChartErrorBoundaryProps, IChartErrorBoundaryState> {
  public state: IChartErrorBoundaryState = { error: null };

  public static getDerivedStateFromError(error: Error): IChartErrorBoundaryState {
    return { error };
  }

  public componentDidUpdate(prev: IChartErrorBoundaryProps): void {
    if (this.state.error && prev.resetKey !== this.props.resetKey) {
      this.setState({ error: null });
    }
  }

  public render(): React.ReactNode {
    if (this.state.error) {
      return (
        <div className={styles.errorMessage} role="alert">
          {fmt(strings.ChartRenderErrorLabel, this.state.error.message)}
        </div>
      );
    }
    return this.props.children;
  }
}
