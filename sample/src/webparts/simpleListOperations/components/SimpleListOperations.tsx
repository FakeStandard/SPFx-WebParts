import * as React from "react";
import styles from "./SimpleListOperations.module.scss";
import { ISimpleListOperationsProps } from "./ISimpleListOperationsProps";
import {
  IListItem,
  ISimpleListOperationsStates,
} from "./ISimpleListOperationsWebPartStates";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  TextField,
  DefaultButton,
  PrimaryButton,
  Stack,
  IStackTokens,
  IIconProps,
} from "office-ui-fabric-react/lib/";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
// import { autobind } from "office-ui-fabric-react/lib/Utilities";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

const stackTokens: IStackTokens = { childrenGap: 40 };
const DelIcon: IIconProps = { iconName: "Delete" };
const ClearIcon: IIconProps = { iconName: "Clear" };
const AddIcon: IIconProps = { iconName: "Add" };

export default class SimpleListOperations extends React.Component<
  ISimpleListOperationsProps,
  ISimpleListOperationsStates
> {
  public constructor(prop: ISimpleListOperationsProps) {
    super(prop);
    this.state = {
      addText: "",
      updateText: [],
    };

    if (Environment.type === EnvironmentType.SharePoint) {
      this._getListItems();
    } else if (Environment.type === EnvironmentType.Local) {
      // return (<div>Whoops! you are using local host...</div>);
    }
  }

  public render(): React.ReactElement<ISimpleListOperationsProps> {
    return (
      <div className={styles.simpleListOperations}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {this.state.updateText.map((row, index) => (
                <Stack horizontal tokens={stackTokens}>
                  <TextField
                    label="Title"
                    underlined
                    value={row.title}
                    onChange={(e, textval) => {
                      // this.setState({ addText: textval });
                      row.title = textval;
                    }}
                  ></TextField>
                  <PrimaryButton
                    text="Update"
                    onClick={() => this._updateClicked(row)}
                  />
                  <DefaultButton
                    text="Delete"
                    onClick={() => this._deleteClicked(row)}
                    iconProps={DelIcon}
                  />
                </Stack>
              ))}

              <br></br>
              <hr></hr>
              <label>Create new item</label>
              <Stack horizontal tokens={stackTokens}>
                <TextField
                  label="Title"
                  underlined
                  value={this.state.addText}
                  onChange={(e, textval) => this.setState({ addText: textval })}
                ></TextField>
                <PrimaryButton
                  text="Save"
                  onClick={this._addClicked}
                  iconProps={AddIcon}
                />
                <DefaultButton
                  text="Clear"
                  onClick={this._clearClicked}
                  iconProps={ClearIcon}
                />
              </Stack>
            </div>
          </div>
        </div>
      </div>
    );
  }

  async _getListItems() {
    const allItems: any[] = await sp.web.lists
      .getByTitle("Colors")
      .items.getAll();
    let items: IListItem[] = [];
    allItems.forEach((element) => {
      items.push({ id: element.Id, title: element.Title });
    });
    this.setState({ updateText: items });
  }

  // @autobind
  _updateClicked = async (row: IListItem) => {
    await sp.web.lists.getByTitle("Colors").items.getById(row.id).update({
      Title: row.title,
    });
  };

  // @autobind
  _deleteClicked = async (row: IListItem) => {
    await sp.web.lists.getByTitle("Colors").items.getById(row.id).recycle();
    this._getListItems();
  };

  // @autobind
  _addClicked = async () => {
    await sp.web.lists.getByTitle("Colors").items.add({
      Title: this.state.addText,
    });
    this.setState({ addText: "" });
    this._getListItems();
  };

  // @autobind
  private _clearClicked = (): void => {
    this.setState({ addText: "" });
  };
}
