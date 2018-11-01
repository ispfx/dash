import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps,
} from '@microsoft/sp-webpart-base';
import { ColorPalette, IColorPaletteProps } from './components/ColorPalette';

export interface IPropertyPaneColorPaletteProps {
  label: string;
  key: string;
  colors: string[];
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  disabled?: boolean;
}

export interface IPropertyPaneColorPaletteInternalProps extends IPropertyPaneColorPaletteProps, IPropertyPaneCustomFieldProps {}

export class PropertyPaneColorPalette implements IPropertyPaneField<IPropertyPaneColorPaletteProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneColorPaletteInternalProps;
  private elem: HTMLElement;

  constructor(targetProperty: string, properties: IPropertyPaneColorPaletteProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      label: properties.label,
      key: properties.key,
      colors: properties.colors,
      onPropertyChange: properties.onPropertyChange,
      disabled: properties.disabled,
      onRender: this.onRender.bind(this),
    };
  }

  public render(): void {
    if (!this.elem) return;
    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) this.elem = elem;

    const element: React.ReactElement<{}> = React.createElement(ColorPalette, {
      colors: this.properties.colors,
      disabled: this.properties.disabled,
      onChanged: this.onChanged.bind(this),
      key: this.properties.key,
    });

    ReactDom.render(element, elem);
  }

  private onChanged(colors: string[]): void {
    this.properties.onPropertyChange(this.targetProperty, this.properties.colors, colors);
  }
}
