import { Web } from "@pnp/sp/presets/all";
import { autobind, TextField, Text, MessageBar, MessageBarType } from "office-ui-fabric-react";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { IStackProps, Stack } from "office-ui-fabric-react/lib/Stack";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
import * as React from "react";
import styles from "./App.module.scss";
import { IAppProps, IAppState } from "./IAppProps";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 500 } }
};

const options: IDropdownOption[] = [
  { key: "Finance Template", text: "Finance Template" },
  { key: "HR Template", text: "HR Template" },
  { key: "IT Temaplate", text: "IT Temaplate" }
];

export default class App extends React.Component<IAppProps, IAppState> {
  constructor(props: IAppProps, state: IAppState) {
    super(props);
    this.state = {
      addUsers: [],
      siteTitle: "",
      siteTitle_Error: false,
      description: "",
      tenentURL: "",
      template: "",
      sucessMessage:false,
      errorMessage:false,
    };
  }

  public componentDidMount(): void {
    this.setState({
      tenentURL:
        "https://sohodragon.sharepoint.com/sites/SharePointSaturday-MegaDemo/"
    });
  }

  @autobind
  private addSelectedUsers(): void {
    let ids = []; let users = this.state.addUsers;
    for (let index = 0; index < users.length; index++) {
      ids.push(users[index]["id"]);
    }
    console.log(ids);
    let newWeb = Web(this.state.tenentURL);
    newWeb.lists
      .getByTitle("Site Request")
      .items.add({
        Title: this.state.siteTitle,
      URL: this.state.siteTitle.replace(/\s/g, ""),
      AdminsId: {
        "results":ids
      },
      Description:
        this.state.description !== ""
          ? this.state.description.toString()
          : "",
      Template: this.state.template
      })
      .then(i => {
        this.setState({
          sucessMessage:!this.state.sucessMessage,
          errorMessage:false,
        })
      })
      .catch(e => {
        console.log(e.toString());
        this.setState({
          errorMessage:!this.state.errorMessage,
          sucessMessage:false
        })
      });
  }

  private _getPeoplePickerItems = (items: any[]) => {
    this.setState({
      addUsers: items
    });
  };

  public render(): React.ReactElement<IAppProps> {
    return (
      <div className={styles.app}>
        <div className={styles.container}>
          <div className={styles.row}>
            <Stack>
              <Stack {...columnProps} style={{ border: "1px solid #cecece" }}>
                <Stack
                  style={{
                    background: "#cecece",
                    padding: "15px",
                    borderBottom: "1px solid black"
                  }}
                >
                  <Text
                    variant={"large"}
                    style={{
                      color: "black",
                      textAlign: "center",
                      fontSize: "28px",
                      fontWeight: "bold"
                    }}
                  >
                    Site Request
                  </Text>
                  <Text
                    variant={"medium"}
                    style={{ color: "black", textAlign: "center" }}
                  >
                    Please Enter the Details to request new Site
                  </Text>
                </Stack>
                <Stack style={{ padding: "0px 15px" }}>
                  <TextField
                    label="New Site Title"
                    value={this.state.siteTitle}
                    ariaLabel="New Site Title"
                    required
                    errorMessage={
                      this.state.siteTitle_Error !== true
                        ? null
                        : "Title Cannot be Empty"
                    }
                    onChange={(
                      event: React.FormEvent<
                        HTMLInputElement | HTMLTextAreaElement
                      >,
                      newValue: string
                    ) => {
                      if (newValue !== "") {
                        this.setState({
                          siteTitle: newValue
                        });
                      } else {
                        this.setState({
                          siteTitle_Error: !this.state.siteTitle_Error
                        });
                      }
                    }}
                  />
                  <TextField
                    label="Description"
                    value={this.state.description}
                  />
                  <Dropdown
                    placeholder="Select an option"
                    label="Please Select Template"
                    options={options}
                    //value={this.state.template}
                    onChange={(
                      event: React.FormEvent<HTMLDivElement>,
                      item: IDropdownOption
                    ) => {
                      this.setState({ template: item.text });
                    }}
                  />
                  <PeoplePicker
                    context={this.props.context}
                    titleText="Please Enter Admin Details"
                    personSelectionLimit={2}
                    groupName={""} // Leave this blank in case you want to filter from all users
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getPeoplePickerItems}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  />
                  <br />
                  <PrimaryButton
                    data-automation-id="addSelectedUsers"
                    title="Add Selected Users"
                    style={{ float: "right" }}
                    onClick={this.addSelectedUsers}
                  >
                    Request Site
                  </PrimaryButton>
                  <br />
                  {this.state.sucessMessage !== false ? <React.Fragment>
                    <MessageBar
                    messageBarType={MessageBarType.success}
                    isMultiline={false}
                  >
                    Request Completed Successfully.
                  </MessageBar>
                  </React.Fragment>:null}
                  {this.state.errorMessage !== false ? <React.Fragment>
                    <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                  >
                    Something Went Wrong Please Contact IT Admin.
                  </MessageBar>
                  </React.Fragment>:null}
                </Stack>
              </Stack>
            </Stack>
          </div>
        </div>
      </div>
    );
  }
}

// import * as React from "react";
// import styles from "./App.module.scss";
// import { IAppProps, IAppState } from "./IAppProps";
// import { Web } from "@pnp/sp/presets/all";
// //import { Web } from "@pnp";
// import {
//   PeoplePicker,
//   PrincipalType
// } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import {
//   IButtonProps,
//   DefaultButton,
//   PrimaryButton
// } from "office-ui-fabric-react/lib/Button";
// import { autobind, Text, ITextStyles } from "office-ui-fabric-react";
// import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
// import {
//   TextField,
//   MaskedTextField
// } from "office-ui-fabric-react/lib/TextField";
// import {
//   Card,
//   ICardTokens,
//   ICardSectionStyles,
//   ICardSectionTokens
// } from "@uifabric/react-cards";
