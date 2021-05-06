import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'VisioForTheWebWebPartStrings';
import VisioForTheWeb from './components/VisioForTheWeb';
import { IVisioForTheWebProps } from './components/IVisioForTheWebProps';

export interface IVisioForTheWebWebPartProps {
  visiofileurl: string;
  visioForTheWebObject: VisioForTheWebObject;
  shapeName: string;
  bHighLight: boolean;
}

import 'officejs';
import { VisioForTheWebObject } from "../../shared/VisioForTheWebObject";

export default class VisioForTheWebWebPart extends BaseClientSideWebPart<IVisioForTheWebWebPartProps> {
  private visioForTheWebObject: VisioForTheWebObject;
  private shapeName: string;

  public onInit(): Promise<void> {
      this.visioForTheWebObject = new VisioForTheWebObject();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IVisioForTheWebProps> = React.createElement(
      VisioForTheWeb,
      {
        visiofileurl: this.properties.visiofileurl,
        visioForTheWebObject: this.visioForTheWebObject,
        shapeName: this.properties.shapeName,
        bHighLight:this.properties.bHighLight
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected HighlightToggleClick(oldVal: any): any {
    this.properties.bHighLight = !this.properties.bHighLight;
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
                PropertyPaneTextField('visiofileurl', {
                  label: strings.VisioFileUrlFieldLabel
                }),
                PropertyPaneTextField('shapeName', {
                  label: strings.ShapeNameLabel
                }),
                PropertyPaneButton('highlightShape', {
                  text: 'Highlight shape toggle',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.HighlightToggleClick.bind(this)
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
