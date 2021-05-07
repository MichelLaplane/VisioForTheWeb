import * as React from 'react';
import styles from './VisioForTheWeb.module.scss';
import { IVisioForTheWebProps } from './IVisioForTheWebProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class VisioForTheWeb extends React.Component<IVisioForTheWebProps, {}> {
  persistentValue:string;

  constructor(props: IVisioForTheWebProps) {
    super(props);

    // set delegate functions that will be used to pass the values from the Visio service to the component
    // this.props.visioService.onSelectionChanged = this._onSelectionChanged;
  }

  public render(): React.ReactElement<IVisioForTheWebProps> {
    return (
      <div className={styles.visioForTheWeb}>
        <div id='iframeHost' ></div>
      </div>
    );
  }

  public componentDidMount() {
        if (this.props.visiofileurl) {
      this.props.visioForTheWebObject.load(this.props.visiofileurl);
    }
  }

  public async componentDidUpdate(prevProps: IVisioForTheWebProps) {
    if (this.props.visiofileurl && this.props.visiofileurl !== prevProps.visiofileurl) {
      this.props.visioForTheWebObject.load(this.props.visiofileurl);
    }
    if ((this.props.bHighLight !== prevProps.bHighLight) || (this.props.shapeName !== prevProps.shapeName)) {
      this.props.visioForTheWebObject.highlightShape(this.props.shapeName, this.props.bHighLight);
    }
    // this.props.visioForTheWebObject.MyFunction();
  }


}
