import * as React from 'react';
import styles from './PeoplepickerSample.module.scss';
import { IPeoplepickerSampleProps } from './IPeoplepickerSampleProps';
import { IPeoplepickerSampleStates } from './IPeoplepickerSampleStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { IStackProps, MessageBarType } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DefaultButton, MessageBar, Stack, TextField, ThemeSettingName } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

const verticalStackProps: IStackProps = {
  styles: { root: { overflow: 'hidden', width: '100%' } },
  tokens: { childrenGap: 20 }
}

export default class PeoplepickerSample extends React.Component<IPeoplepickerSampleProps, IPeoplepickerSampleStates> {
  constructor(props: IPeoplepickerSampleProps, state: IPeoplepickerSampleStates) {
    super(props);

    sp.setup({
      spfxContext: this.props.context
    });

    this.state = {
      title: '',
      users: [],
      managers: [],
      showMessageBar: false
    }
  }

  _onChangedUser = (users: any) => {
    let items = [];
    users.forEach(res => {
      items.push(res.id);
    })

    this.setState({ users: items });
  }

  _onChangedManager = (managers: any) => {
    let items = [];
    managers.forEach(res => {
      items.push(res.id);
    })

    this.setState({ managers: items });
  }

  // _onChangedTitle = (title: any) => {
  //   this.setState({ title: title });
  // }

  _createItem = async () => {
    try {
      const item = await sp.web.lists.getByTitle("Users").items.get();
      await sp.web.lists.getByTitle("Users").items.add({
        // Title: this.state.title,
        UserId: { results: this.state.users },
        ManagerId: { results: this.state.managers }
      });

      this.setState({
        message: `Item ${this.state.title} created successfully!`,
        showMessageBar: true,
        messageType: MessageBarType.success
      });
    } catch (error) {
      this.setState({
        message: `Item ${this.state.title} creation failed with error: ${error}}`,
        showMessageBar: true,
        messageType: MessageBarType.error
      })
    }
  }

  public render(): React.ReactElement<IPeoplepickerSampleProps> {
    return (
      <div className={styles.peoplepickerSample}>
        {this.state.showMessageBar ?
          <div className="form-group">
            <Stack {...verticalStackProps}>
              <MessageBar messageBarType={this.state.messageType}>{this.state.message}</MessageBar>
            </Stack>
          </div>
          : null}
        <br />
        {/* <TextField label='Title' required onChange={this._onChangedTitle} /> */}
        <br />
        <PeoplePicker
          context={this.props.context}
          titleText="User"
          personSelectionLimit={1}
          showtooltip={true}
          required={true}
          disabled={false}
          onChange={this._onChangedUser}
          ensureUser={true}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000} />
        <PeoplePicker
          context={this.props.context}
          titleText="Manager"
          personSelectionLimit={3}
          showtooltip={true}
          required={true}
          disabled={false}
          onChange={this._onChangedManager}
          ensureUser={true}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000} />
        <br />
        <DefaultButton text="Submit" onClick={this._createItem} />
      </div>
    );
  }
}
