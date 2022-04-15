import * as React from 'react';
import styles from './TemperatureConversion.module.scss';
import { ITemperatureConversionProps } from './ITemperatureConversionProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class TemperatureConversion extends React.Component<ITemperatureConversionProps, {}> {
  public render(): React.ReactElement<ITemperatureConversionProps> {
    return (
      <div className={ styles.temperatureConversion }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
