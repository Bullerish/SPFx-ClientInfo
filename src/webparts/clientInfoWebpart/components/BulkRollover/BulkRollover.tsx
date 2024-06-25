import * as React from 'react';
import { useState, useEffect } from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Text } from "office-ui-fabric-react/lib/Text";
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";


import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { IItemAddResult } from "@pnp/sp/items";
import {
  DateTimePicker,
  DateConvention,
  TimeConvention,
} from "@pnp/spfx-controls-react/lib/DateTimePicker";
import {
  DatePicker,
  DayOfWeek,
  IDatePickerStrings,
} from "office-ui-fabric-react/lib/DatePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";


import {
  ListView,
  IViewField,
  SelectionMode,
  GroupOrder,
  IGrouping,
} from "@pnp/spfx-controls-react/lib/ListView";



export interface IDatePickerFormatExampleState {
  firstDayOfWeek?: DayOfWeek;
  value?: Date | null;
}


const BulkRollover = ({
  spContext,
  isBulkRolloverOpen,
  onBulkRolloverModalHide,
}): React.ReactElement => {
  // all state variables
  const [team, setTeam] = useState<string>("");
  const [isDataLoaded, setIsDataLoaded] = useState<boolean>(true);
  const [items, setItems] = useState([
    {
      EngagementName: "Engagement 1",
      Title: "Matter 1",
      TemplateType: "Template 1",
      SiteOwner: "Owner 1",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 2",
      Title: "Matter 2",
      TemplateType: "Template 2",
      SiteOwner: "Owner 2",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 3",
      Title: "Matter 3",
      TemplateType: "Template 3",
      SiteOwner: "Owner 3",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 4",
      Title: "Matter 4",
      TemplateType: "Template 4",
      SiteOwner: "Owner 4",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 5",
      Title: "Matter 5",
      TemplateType: "Template 5",
      SiteOwner: "Owner 5",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 6",
      Title: "Matter 6",
      TemplateType: "Template 6",
      SiteOwner: "Owner 6",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 7",
      Title: "Matter 7",
      TemplateType: "Template 7",
      SiteOwner: "Owner 7",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 8",
      Title: "Matter 8",
      TemplateType: "Template 8",
      SiteOwner: "Owner 8",
      PortalExpiration: "12/31/2021",
    },
    {
      EngagementName: "Engagement 9",
      Title: "Matter 9",
      TemplateType: "Template 9",
      SiteOwner: "Owner 9",
      PortalExpiration: "12/31/2021",
    },
  ]);
  const [portalSelected, setPortalSelected] = useState([]);


  // store the current site's absolute URL (should be a client site URL)
  const advAbsoluteUrl = spContext._pageContext._web.absoluteUrl;


  // runs after a radio button is selected for Team type
  useEffect(() => {
    console.log("team selected::", team);
  }, [team]);


  // function to reset all state variables back to their initial values
  const resetState = () => {
    alert("resetState fired::");
    onBulkRolloverModalHide(false);
  };

  // onChange event handler to capture the selected team value
  const onTeamChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    console.log("onTeamChange fired::");
    console.log(option.key);
    setTeam(option.key);
  };

  // function to create the initial date for the DatePicker component
  const createDate18MonthsFromNow = (): Date => {
    console.log("createDate18MonthsFromNow fired::");
    const date = new Date();
    date.setMonth(date.getMonth() + 18);
    return date;
  };

  // function to format the date for the DatePicker component
  const onFormatDate = (date: Date): string => {
    return (
      (date.getMonth() + 1) +
      "/" +
      date.getDate() +
      "/" +
      (date.getFullYear() % 100)
    );
  };

  // function to capture the selected date from the DatePicker component
  const onSelectDate = (date: Date | null | undefined): void => {
    console.log("onSelectDate fired::");
    console.log(date);
  };

  const getPeoplePickerItems = (items: any[]) => {
    const currSite = Web(GlobalValues.HubSiteURL);
    let getSelectedUsers = [];
    let getusersEmails = [];
    for (let item in items) {
      getSelectedUsers.push(items[item].text);
      getusersEmails.push(items[item].secondaryText);
    }
    items.forEach((e) => {
      currSite.siteUsers
        .getByLoginName(e.loginName)
        .get()
        .then((user) => {
          // this.setState({
          //   addusers: getSelectedUsers,
          //   addusersID: user.Id,
          //   emailaddress: getusersEmails,
          // });
          console.log("logging user info:: ", user);
        });
    });
  }

  const validateSiteOwner = (items: any[]) => {
    // show error message if this is a guest user
    if (items.length > 0) {
      let userEmail = items[0].secondaryText.toLowerCase();
      if (
        userEmail.indexOf("cohnreznick.com") == -1 &&
        userEmail.indexOf("cohnreznickdev") == -1
      ) {
        // this is a guest user, do not validate
        // this.setState({ addusers: [] });
      } else {
        getPeoplePickerItems(items);
      }
    } else {
      // this.setState({ addusers: [] });
    }
  }

  // define columns/viewfields so the ListView component knows what to render
  const viewFields: IViewField[] = [
    {
      name: "EngagementName",
      displayName: "Engagement Name",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
    },
    {
      name: "Title",
      displayName: "Matter #",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
    },
    {
      name: "TemplateType",
      displayName: "Template Type",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
    {
      name: "SiteOwner",
      displayName: "Site Owner",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
      render: (rowItem, index, column) => {


        return (
        <div>
          <PeoplePicker
            context={spContext}
            showtooltip={false}
            required={true}
            onChange={(items) => validateSiteOwner(items)}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            personSelectionLimit={1}
            placeholder="Enter name or email"
            // defaultSelectedUsers={this.state.addusers}
          />
        </div>
        );
      },
    },
    {
      name: "PortalExpiration",
      displayName: "Expiration Date",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
      render: (rowItem, index, column) => {
        // console.log("logging rowItem info:: ", rowItem);
        // console.log("logging index info:: ", index);

        return (
          <div>
            <DatePicker
              // label="DateTime Picker - 24h"
              // dateConvention={DateConvention.DateTime}
              // timeConvention={TimeConvention.Hours24}
              allowTextInput={false}
              value={createDate18MonthsFromNow()}
              initialPickerDate={new Date()}
              onSelectDate={onSelectDate}
              formatDate={onFormatDate}
            />
          </div>
        );
      },
    },
  ];

  return (
    <>
      <Dialog
        hidden={!isBulkRolloverOpen}
        onDismiss={resetState}
        minWidth={900}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Bulk Subportal Rollover",
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
          className: styles.bulkRollover,
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <span className={styles.guidanceText}>
          Choose a team to see WF portals that are available for rollover
        </span>

        {/* Team select choice group */}
        <div>
          <ChoiceGroup
            className={styles.innerChoice}
            // defaultSelectedKey="B"
            label="Team"
            required={true}
            options={[
              {
                key: "assurance",
                text: "Assurance",
              },
              {
                key: "tax",
                text: "Tax",
              },
            ]}
            onChange={onTeamChange}
          />
        </div>

        {/* ListView component to display list of portals */}
        {isDataLoaded ? (
          <>
            <span className={styles.guidanceText}>
              Select engagements below to bulk rollover. No permissions will be rolled over to the new portals.
            </span>

            <ListView
              items={items}
              viewFields={viewFields}
              // iconFieldName="FileRef"
              compact={true}
              selectionMode={SelectionMode.multiple}
              selection={(selectionItem) => setPortalSelected(selectionItem)}
              // defaultSelection={defaultSelectedFromScreen2}
              showFilter={false}
              key="engagementPortals"
            />
          </>
        ) : (
          <Spinner
            size={SpinnerSize.large}
            label="Loading Engagement Portal Data...this could take some time depending on the number of portals you have access to."
          />
        )}

        {/* Dialog footer to hold buttons */}
        <DialogFooter>
          <div className={styles.dialogFooterButtonContainer}>
            <DefaultButton
              className={styles.defaultButton}
              onClick={resetState}
              text="Cancel"
            />
            <PrimaryButton
              className={styles.primaryButton}
              // onClick={() => setConfirmDialogHidden(false)}
              text="Next"
              // disabled={
              //   fullName !== "" &&
              //   fullName !== null &&
              //   jobRoleItem.key &&
              //   jobFunctionItem.key
              //     ? false
              //     : true
              // }
            />
          </div>
        </DialogFooter>
      </Dialog>
      {/* confirm dialog component. Modal/dialog window will open */}
      {/* <ConfirmDialog
        confirmDialogHidden={confirmDialogHidden}
        onSetConfirmDialogHidden={onSetConfirmDialogHidden}
        onConfirmSubmission={submitUserProfileInfo}
      /> */}
    </>
  );
};

export default BulkRollover;
