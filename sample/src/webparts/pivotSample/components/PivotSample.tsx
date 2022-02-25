import * as React from "react";
import styles from "./PivotSample.module.scss";
import { IPivotSampleProps } from "./IPivotSampleProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  ILabelStyles,
  IStyleSet,
  Label,
  Pivot,
  PivotItem,
} from "@fluentui/react";
import PivotComponent from "./PivotComponent";

export interface IPivotItem {
  key: string;
  text: string;
}

const ISampleItems: IPivotItem[] = [
  { key: "0", text: "Dynamic Item #1" },
  { key: "1", text: "Dynamic Item #2" },
  { key: "2", text: "Dynamic Item #3" },
  { key: "3", text: "Dynamic Item #4" },
];

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 },
};

export default class PivotSample extends React.Component<
  IPivotSampleProps,
  {}
> {
  public render(): React.ReactElement<IPivotSampleProps> {
    return (
      <div>
        {/* Simple */}
        <Pivot>
          <PivotItem headerText="Tab 1">
            <Label styles={labelStyles}>Pivot #1 content</Label>
          </PivotItem>
          <PivotItem headerText="Tab 2">
            <Label styles={labelStyles}>Pivot #2 content</Label>
          </PivotItem>
          <PivotItem headerText="Tab 3">
            <Label styles={labelStyles}>Pivot #3 content</Label>
          </PivotItem>
        </Pivot>
        <hr />

        {/* Dynamic */}
        <Pivot linkSize="large">
          {ISampleItems.map((item: IPivotItem) => {
            return (
              <PivotItem headerText={item.text} key={item.key}>
                <Label styles={labelStyles}>
                  This is the {item.text} with key {item.key}
                </Label>
              </PivotItem>
            );
          })}
        </Pivot>
        <hr />
        
        {/* Import Component*/}
        <PivotComponent />
      </div>
    );
  }
}
