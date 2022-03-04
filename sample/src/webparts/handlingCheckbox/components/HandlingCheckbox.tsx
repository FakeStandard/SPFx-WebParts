import * as React from "react";
import styles from "./HandlingCheckbox.module.scss";
import { IHandlingCheckboxProps } from "./IHandlingCheckboxProps";
import { IHandlingCheckboxStates } from "./IHandlingCheckboxStates";
import { escape } from "@microsoft/sp-lodash-subset";
import { Checkbox, Label, Stack } from "@fluentui/react";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DefaultButton } from "@microsoft/office-ui-fabric-react-bundle";

const options: string[] = [
  "White",
  "Black",
  "Red",
  "Blue",
  "Green",
  "Yellow",
  "Grey",
  "Orange",
  "Purple",
  "Pink",
  "Brown",
];

export default class HandlingCheckbox extends React.Component<
  IHandlingCheckboxProps,
  IHandlingCheckboxStates
> {
  constructor(props) {
    super(props);
    this.state = {
      basicText: "",
      dynamicText: "",
      colorArr: [],
      message: "",
    };
  }

  public render(): React.ReactElement<IHandlingCheckboxProps> {
    return (
      <div className={styles.handlingCheckbox}>
        <Checkbox label="Choices" onChange={this._onChange} />
        <Label>{this.state.basicText}</Label>
        <br />

        <Stack tokens={{ childrenGap: 7 }}>
          {options.map((item) => (
            <Checkbox
              label={item}
              title={item}
              onChange={this._onChangeColor}
            />
          ))}
        </Stack>
        <Label>{this.state.dynamicText}</Label>
        <DefaultButton text="Save" onClick={this._onClick} />
        <Label>{this.state.message}</Label>
        <br />
      </div>
    );
  }

  private _onChange = (
    ev: React.FormEvent<HTMLElement>,
    isChecked: boolean
  ) => {
    let text = `The option has been changed to ${isChecked}.`;
    this.setState({ basicText: text });
  };

  private _onChangeColor = (ev, isChecked) => {
    let text = `The option ${ev.currentTarget.title} has been changed to ${isChecked}`;
    this.setState({ dynamicText: text });

    this.getColor();

    let colorArr: string[] = this.state.colorArr;
    if (isChecked) colorArr.push(ev.currentTarget.title);
    else {
      colorArr.forEach((i, index) => {
        if (i === ev.currentTarget.title) {
          colorArr.splice(index, 1);
        }
      });
    }

    this.setState({ colorArr: colorArr, message: "" });
  };

  private _onClick = () => {
    try {
      sp.web.lists.getByTitle("Colors").items.add({
        Title: "polar",
        Color: { results: this.state.colorArr },
      });

      this.setState({ message: "success" });
    } catch {
      this.setState({ message: "error!" });
    }
  };

  private getColor = () => {
    sp.web.lists
      .getByTitle("Colors")
      .items.get()
      .then((result) => {
        console.log(result);
      });
  };
}
