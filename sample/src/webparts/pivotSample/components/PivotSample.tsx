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
      </div>
    );
  }
}
