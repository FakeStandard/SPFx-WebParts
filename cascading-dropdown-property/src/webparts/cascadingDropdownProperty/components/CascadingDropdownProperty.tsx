import * as React from "react";
import styles from "./CascadingDropdownProperty.module.scss";
import { ICascadingDropdownPropertyProps } from "./ICascadingDropdownPropertyProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class CascadingDropdownProperty extends React.Component<
  ICascadingDropdownPropertyProps,
  {}
> {
  public render(): React.ReactElement<ICascadingDropdownPropertyProps> {
    return (
      <div className={styles.cascadingDropdownProperty}>
        <div className={styles.container}>
          <div
            className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}
          >
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">
                Welcome to SharePoint!
              </span>
              <p className="ms-font-l ms-fontColor-white">
                Customize SharePoint experiences using web parts.
              </p>
              <p className="ms-font-l ms-fontColor-white">
                {escape(this.props.listName)}
              </p>
              <p className="ms-font-l ms-fontColor-white">
                {escape(this.props.itemName)}
              </p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
