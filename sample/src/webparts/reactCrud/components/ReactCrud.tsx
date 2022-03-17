import * as React from "react";
import styles from "./ReactCrud.module.scss";
import { IReactCrudProps } from "./IReactCrudProps";
import { IReactCrudStates } from "./IReactCrudStates";
import { escape } from "@microsoft/sp-lodash-subset";
import { Label, PrimaryButton } from "@microsoft/office-ui-fabric-react-bundle";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from "office-ui-fabric-react";
import {
  DatePicker,
  IDatePickerStrings,
} from "office-ui-fabric-react/lib/DatePicker";
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Stack } from "@fluentui/react";

const DatePickerStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],
  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ],
  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],
  shortDays: ["S", "M", "T", "W", "T", "F", "S"],
  goToToday: "Go to today",
  prevMonthAriaLabel: "Go to previous month",
  nextMonthAriaLabel: "Go to next month",
  prevYearAriaLabel: "Go to previous year",
  nextYearAriaLabel: "Go to next year",
  invalidInputErrorMessage: "Invalid date format.",
};

const FormatDate = (date): string => {
  console.log(date);
  var date1 = new Date(date);
  var year = date1.getFullYear();
  var month = (1 + date1.getMonth()).toString();
  month = month.length > 1 ? month : "0" + month;
  var day = date1.getDate().toString();
  day = day.length > 1 ? day : "0" + day;
  return month + "/" + day + "/" + year;
};

export default class ReactCrud extends React.Component<
  IReactCrudProps,
  IReactCrudStates
> {
  constructor(props) {
    super(props);
    this.state = {
      HTML: [],
      Items: [],
      ID: 0,
      EmployeeName: "",
      EmployeeNameId: 0,
      HireDate: null,
      JobDescription: "",
    };
  }

  public componentWillMount(): void {
    console.log("componentWillMount");
    // before component
  }

  public componentDidMount = async () => {
    console.log("componentDidMount");
    await this.fetchData();
    // after component
  };

  public componentDidUpdate(prevProps, prevStates): void {
    // when component change
    console.log("componentDidUpdate");
  }

  // public async componentDidMount() {
  // }

  public async fetchData() {
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists
      .getByTitle("EmployeeDetails")
      .items.select("*", "EmployeeName/Title")
      .expand("EmployeeName")
      .getAll();

    this.setState({ Items: items });
    let html = await this.getHTML(items);
    this.setState({ HTML: html });
  }

  public findData = (id): void => {
    var itemID = id;
    var allitems = this.state.Items;
    var allitemsLength = allitems.length;
    if (allitemsLength > 0) {
      for (var i = 0; i < allitemsLength; i++) {
        if (itemID == allitems[i].Id) {
          this.setState({
            ID: itemID,
            EmployeeName: allitems[i].EmployeeName.Title,
            EmployeeNameId: allitems[i].EmployeeNameId,
            HireDate: new Date(allitems[i].HireDate),
            JobDescription: allitems[i].JobDescription,
          });
        }
      }
    }
  };

  public async getHTML(items) {
    var tabledata = (
      <table>
        <thead>
          <tr>
            <th>Employee Name</th>
            <th>Hire Date</th>
            <th>Job Description</th>
          </tr>
        </thead>
        <tbody>
          {items &&
            items.map((item, i) => {
              return [
                <tr key={i} onClick={() => this.findData(item.ID)}>
                  <td>{item.EmployeeName.Title}</td>
                  <td>{FormatDate(item.HireDate)}</td>
                  <td>{item.JobDescription}</td>
                </tr>,
              ];
            })}
        </tbody>
      </table>
    );
    return await tabledata;
  }

  public render(): React.ReactElement<IReactCrudProps> {
    return (
      <div className={styles.reactCrud}>
        <h1>CRUD Operations With ReactJS</h1>
        {this.state.HTML}
        <br />
        <br />
        <div>
          <Stack tokens={{ childrenGap: 10 }} horizontal>
            <PrimaryButton text="Create" onClick={() => this.SaveData()} />
            <PrimaryButton text="Update" onClick={() => this.UpdateData()} />
            <PrimaryButton text="Delete" onClick={() => this.DeleteData()} />
          </Stack>
          <div>
            <form>
              <div>
                <Label>Employee Name</Label>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  required={false}
                  onChange={this._getPeoplePickerItems}
                  defaultSelectedUsers={[
                    this.state.EmployeeName ? this.state.EmployeeName : "",
                  ]}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                ></PeoplePicker>
              </div>
              <div>
                <Label>Hire Date</Label>
                <DatePicker
                  maxDate={new Date()}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  value={this.state.HireDate}
                  onSelectDate={(e) => {
                    this.setState({ HireDate: e });
                  }}
                  ariaLabel="Select a date"
                  formatDate={FormatDate}
                />
              </div>
              <div>
                <Label>Job Description</Label>
                <TextField
                  value={this.state.JobDescription}
                  multiline
                  onChange={(e, text) => {
                    this.setState({ JobDescription: text });
                  }}
                />
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  private async SaveData() {
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("EmployeeDetails")
      .items.add({
        EmployeeNameId: this.state.EmployeeNameId,
        HireDate: new Date(this.state.HireDate),
        JobDescription: this.state.JobDescription,
      })
      .then((i) => {
        console.log(i);
      });
    alert("Created Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "" });
    this.fetchData();
  }

  private async UpdateData() {
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("EmployeeDetails")
      .items.getById(this.state.ID)
      .update({
        EmployeeNameId: this.state.EmployeeNameId,
        HireDate: new Date(this.state.HireDate),
        JobDescription: this.state.JobDescription,
      })
      .then((i) => {
        console.log(i);
      });
    alert("Updated Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "" });
    this.fetchData();
  }

  private async DeleteData() {
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("EmployeeDetails")
      .items.getById(this.state.ID)
      .delete()
      .then((i) => {
        console.log(i);
      });
    alert("Deleted Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "" });
    this.fetchData();
  }

  public _getPeoplePickerItems = async (items: any[]) => {
    if (items.length > 0) {
      this.setState({ EmployeeName: items[0].text });
      this.setState({ EmployeeNameId: items[0].id });
    } else {
      //ID=0;
      this.setState({ EmployeeNameId: "" });
      this.setState({ EmployeeName: "" });
    }
  };
}
