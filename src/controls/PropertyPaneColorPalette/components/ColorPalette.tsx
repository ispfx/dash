import * as React from 'react';

export interface IColorPaletteProps {
  colors: string[];
  disabled?: boolean;
  onChanged(colors: string[]): void;
}

export class ColorPalette extends React.Component<IColorPaletteProps> {
  constructor(props: IColorPaletteProps) {
    super(props);

    // Bind methods

  }

  public render(): JSX.Element {
    return (
      <div>
        {this.props.colors.map((color, i) => {
          return (
            <input key={i} type="text" value={color} onChange={event => this.onChanged(event.currentTarget.value, i)} />
          );
        })}
      </div>
    );
  }

  public onChanged(newColor: string, index: number): void {
    const updatedColors = this.props.colors;
    updatedColors[index] = newColor;

    this.props.onChanged(updatedColors);
  }
}
