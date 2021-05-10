import * as React from 'react';
import styles from './VisioForTheWeb.module.scss';
import { IVisioForTheWebProps } from './IVisioForTheWebProps';
import { IVisioForTheWebState } from './IVisioForTheWebState';
import { VisioForTheWebObject } from "../../../shared/VisioForTheWebObject";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { TextField } from 'office-ui-fabric-react/lib/TextField';


export default class VisioForTheWeb extends React.Component<IVisioForTheWebProps, IVisioForTheWebState> {
  private visioForTheWebObject: VisioForTheWebObject;
  private iPrevHighLight: boolean;

  constructor(props: IVisioForTheWebProps) {
    super(props);
    this.visioForTheWebObject = new VisioForTheWebObject();
    this.visioForTheWebObject.onShapeNameEntered = this.onShapeNameEntered.bind(this);
    this.iPrevHighLight = false;
    this.state = {
      iHighLight: false,
      shapeNameFlyout: "",
    };
  }

  public render(): React.ReactElement<IVisioForTheWebProps> {
    return (
      <div className={styles.visioForTheWeb}>
        <div id='iframeHost' ></div>
        <div  >
          <Stack horizontal tokens={{ childrenGap: 20 }} >
            <PrimaryButton text="Highlight shape toggle" onClick={this.HighlightToggleClick.bind(this)} />
            <TextField id='iHighLight' label="Entered shape:" underlined defaultValue={this.state.shapeNameFlyout} onChange={this.onShapeNameEnteredChange} />
          </Stack>
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
    if ((this.props.bHighLight !== prevProps.bHighLight) && (this.props.shapeName != "")) {
      this.visioForTheWebObject.highlightShape(this.props.shapeName, this.props.bHighLight);
    }
    if ((this.state.iHighLight != this.iPrevHighLight) && (this.props.shapeName != "")) {
      this.visioForTheWebObject.highlightShape(this.state.shapeNameFlyout, this.state.iHighLight);
      this.iPrevHighLight = this.state.iHighLight;
    }
  }

  private HighlightToggleClick() {
    if (this.state.iHighLight == true) {
      this.setState({ iHighLight: false });
    }
    else {
      this.setState({ iHighLight: true });
    }
  }

  private onShapeNameEnteredChange = (event) => {
    this.setState({
      shapeNameFlyout: event.target.value,
    });
  };

  private onShapeNameEntered(enteredShapeName: string) {
    this.setState({ shapeNameFlyout: enteredShapeName });
  }

}
