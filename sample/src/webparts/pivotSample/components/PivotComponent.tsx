import * as React from "react";
import { Pivot, PivotItem, Label } from "@fluentui/react";
import { IPivotComponentProps } from "./IPivotComponentProps";

const PivotComponent: React.FunctionComponent<IPivotComponentProps> = (
  props
) => {
  return (
    <Pivot linkSize="large">
      <PivotItem headerText="Page 1">Page 1 content.</PivotItem>
      <PivotItem headerText="Page 2">Page 3 content.</PivotItem>
      <PivotItem headerText="Page 3">Page 3 content.</PivotItem>
      <PivotItem headerText="Page 4">Page 4 content.</PivotItem>
      <PivotItem headerText="Page 5">Page 5 content.</PivotItem>
    </Pivot>
  );
};

export default PivotComponent;
