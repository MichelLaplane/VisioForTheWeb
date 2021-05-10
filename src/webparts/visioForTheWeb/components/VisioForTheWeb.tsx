import * as React from 'react';
import styles from './VisioForTheWeb.module.scss';
import { IVisioForTheWebProps } from './IVisioForTheWebProps';
import { IVisioForTheWebState } from './IVisioForTheWebState';
import { VisioForTheWebObject } from "../../../shared/VisioForTheWebObject";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export default class VisioForTheWeb extends React.Component<IVisioForTheWebProps, IVisioForTheWebState> {
  private visioForTheWebObject: VisioForTheWebObject;

  constructor(props: IVisioForTheWebProps) {
    super(props);
    this.visioForTheWebObject = new VisioForTheWebObject();
    this.state = {
      iHighLight: false,
    };
  }

  public render(): React.ReactElement<IVisioForTheWebProps> {
    return (
      <div className={styles.visioForTheWeb}>
        <div id='iframeHost' ></div>
        <div  >
          <PrimaryButton text="Highlight shape toggle" onClick={this.HighlightToggleClick.bind(this)} />
        </div>
      </div>
    );
  }

  public componentDidMount() {
    if (this.props.visiofileurl) {
      this.visioForTheWebObject.load(this.props.visiofileurl);
    }
  }

  public async componentDidUpdate(prevProps: IVisioForTheWebProps) {
    if (this.props.visiofileurl && this.props.visiofileurl !== prevProps.visiofileurl) {
      this.visioForTheWebObject.load(this.props.visiofileurl);
    }
    if ((this.props.bHighLight !== prevProps.bHighLight) || (this.props.shapeName !== prevProps.shapeName)) {
      this.setState({ iHighLight: this.props.bHighLight });
      this.visioForTheWebObject.highlightShape(this.props.shapeName, this.state.iHighLight);
    }
  }

  private HighlightToggleClick() {
    if (this.state.iHighLight == true) {
      this.setState({ iHighLight: false });
    }
    else {
      this.setState({ iHighLight: true });
    }
    this.visioForTheWebObject.highlightShape(this.props.shapeName, this.state.iHighLight);
  }



}
