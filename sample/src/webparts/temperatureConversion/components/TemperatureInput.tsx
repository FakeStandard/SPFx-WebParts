import * as React from "react";

interface ITemperatureInputProps {
    scale: string;
    temperature: string;
    onTemperatureChange: any;
}

export class TemperatureInput extends React.Component<ITemperatureInputProps, {}> {
    constructor(props) {
        super(props);
        this.handleChange = this.handleChange.bind(this);
    }

    handleChange(e) {
        this.props.onTemperatureChange(e.target.value)
    }

    render(): React.ReactNode {
        const temperature = this.props.temperature;
        const scale = this.props.scale;

        return (
            <fieldset>
                <legend>Enter temperature in {scale}</legend>
                <input value={temperature} onChange={this.handleChange} />
            </fieldset>
        );
    }
}