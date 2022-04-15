import * as React from 'react';
import styles from './TemperatureConversion.module.scss';
import { ITemperatureConversionProps } from './ITemperatureConversionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TemperatureInput } from './TemperatureInput';
import { ITemperatureConversionStates } from './ITemperatureConversionStates';

export default class TemperatureConversion extends React.Component<ITemperatureConversionProps, ITemperatureConversionStates> {
  constructor(props) {
    super(props);

    this.handleCelsiusChange = this.handleCelsiusChange.bind(this);
    this.handleFahrenheitChange = this.handleFahrenheitChange.bind(this);

    this.state = { temperature: '', scale: 'c' };
  }

  handleCelsiusChange(temperature) {
    this.setState({ scale: 'c', temperature });
  }

  handleFahrenheitChange(temperature) {
    this.setState({ scale: 'f', temperature });
  }

  public render(): React.ReactElement<ITemperatureConversionProps> {
    const scale = this.state.scale;
    const temperature = this.state.temperature;
    const celsius = scale === 'f' ? tryConvert(temperature, toCelsius) : temperature;
    const fahrenheit = scale === 'c' ? tryConvert(temperature, toFahrenheit) : temperature;

    return (
      <div className={styles.temperatureConversion}>
        <TemperatureInput scale='Celsius' temperature={celsius} onTemperatureChange={this.handleCelsiusChange} />
        <TemperatureInput scale='Fahrenheit' temperature={fahrenheit} onTemperatureChange={this.handleFahrenheitChange} />
        <BoilingVerdict celsius={parseFloat(celsius)} />
      </div>
    );
  }
}

function toCelsius(fahrenheit) {
  return (fahrenheit - 32) * 5 / 9;
}

function toFahrenheit(celsius) {
  return (celsius * 9 / 5) + 32;
}

function tryConvert(temperature, convert) {
  const input = parseFloat(temperature);
  if (input.toString() == "NaN") {
    return '';
  }
  const output = convert(input);
  const rounded = Math.round(output * 1000) / 1000;
  return rounded.toString();
}

function BoilingVerdict(props) {
  if (props.celsius >= 100) {
    return <p>The water would boil.</p>;
  }
  return <p>The water would not boil.</p>;
}