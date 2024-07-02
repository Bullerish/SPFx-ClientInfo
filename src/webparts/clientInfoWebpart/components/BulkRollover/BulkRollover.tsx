import * as React from 'react';
import { useState, useEffect, useLayoutEffect } from "react";
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
import { getMatterNumbersForClientSite } from './rolloverLogic';



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
      ID: 1,
      EngagementName: "Engagement 1",
      Title: "Matter #1",
      TemplateType: "Template Type 1",
      SiteOwner: "Site Owner 1",
      PortalExpiration: "",
    },
    {
      ID: 2,
      EngagementName: "Engagement 2",
      Title: "Matter #2",
      TemplateType: "Template Type 2",
      SiteOwner: "Site Owner 2",
      PortalExpiration: "",
    },
    {
      ID: 3,
      EngagementName: "Engagement 3",
      Title: "Matter #3",
      TemplateType: "Template Type 3",
      SiteOwner: "Site Owner 3",
      PortalExpiration:
        "Tue Dec 02 2025 00:00:00 GMT-0800 (Pacific Standard Time)",
    },
    {
      ID: 4,
      EngagementName: "Engagement 4",
      Title: "Matter #4",
      TemplateType: "Template Type 4",
      SiteOwner: "Site Owner 4",
      PortalExpiration: "Expiration Date 4",
    },
    {
      ID: 5,
      EngagementName: "Engagement 5",
      Title: "Matter #5",
      TemplateType: "Template Type 5",
      SiteOwner: "Site Owner 5",
      PortalExpiration: "Expiration Date 5",
    },
    {
      ID: 6,
      EngagementName: "Engagement 6",
      Title: "Matter #6",
      TemplateType: "Template Type 6",
      SiteOwner: "Site Owner 6",
      PortalExpiration: "Expiration Date 6",
    },
    {
      ID: 7,
      EngagementName: "Engagement 7",
      Title: "Matter #7",
      TemplateType: "Template Type 7",
      SiteOwner: "Site Owner 7",
      PortalExpiration: "Expiration Date 7",
    },
  ]);
  const [itemsStaged, setItemsStaged] = useState([
    {
      ID: 1,
      EngagementName: "Engagement 1",
      Title: "Matter #1",
      TemplateType: "Template Type 1",
      SiteOwner: "Site Owner 1",
      PortalExpiration: "",
    },
    {
      ID: 2,
      EngagementName: "Engagement 2",
      Title: "Matter #2",
      TemplateType: "Template Type 2",
      SiteOwner: "Site Owner 2",
      PortalExpiration: "",
    },
    {
      ID: 3,
      EngagementName: "Engagement 3",
      Title: "Matter #3",
      TemplateType: "Template Type 3",
      SiteOwner: "Site Owner 3",
      PortalExpiration:
        "Tue Dec 02 2025 00:00:00 GMT-0800 (Pacific Standard Time)",
    },
    {
      ID: 4,
      EngagementName: "Engagement 4",
      Title: "Matter #4",
      TemplateType: "Template Type 4",
      SiteOwner: "Site Owner 4",
      PortalExpiration: "Expiration Date 4",
    },
    {
      ID: 5,
      EngagementName: "Engagement 5",
      Title: "Matter #5",
      TemplateType: "Template Type 5",
      SiteOwner: "Site Owner 5",
      PortalExpiration: "Expiration Date 5",
    },
    {
      ID: 6,
      EngagementName: "Engagement 6",
      Title: "Matter #6",
      TemplateType: "Template Type 6",
      SiteOwner: "Site Owner 6",
      PortalExpiration: "Expiration Date 6",
    },
    {
      ID: 7,
      EngagementName: "Engagement 7",
      Title: "Matter #7",
      TemplateType: "Template Type 7",
      SiteOwner: "Site Owner 7",
      PortalExpiration: "Expiration Date 7",
    },
  ]);
  const [portalSelected, setPortalSelected] = useState([]);
  const [dateSelections, setDateSelections] = useState({});

  // store the current site's absolute URL (should be a client site URL)
  const clientSiteAbsoluteUrl = spContext._pageContext._web.absoluteUrl;
  const clientSiteServerRelativeUrl =
    spContext._pageContext._web.serverRelativeUrl;
  const relativeUrlArr = clientSiteServerRelativeUrl.split("/");
  const clientSiteNumber = relativeUrlArr[relativeUrlArr.length - 1];

  // function to reset all state variables back to their initial values
  const resetState = () => {
    alert("resetState fired::");
    onBulkRolloverModalHide(false);
  };

  // onChange event handler to capture the selected team value
  const onTeamChange = (
    ev: React.FormEvent<HTMLInputElement>,
    option: any
  ): void => {
    console.log("onTeamChange fired::");
    // console.log(option.key);
    setTeam(option.key);
  };


  // function to format the date for the DatePicker component
  const onFormatDate = (date: Date): string => {
    return (
      date.getMonth() +
      1 +
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

  const getPeoplePickerItems = (itemsArr: any[]) => {
    const currSite = Web(GlobalValues.HubSiteURL);
    let getSelectedUsers = [];
    let getusersEmails = [];
    for (let item in itemsArr) {
      getSelectedUsers.push(itemsArr[item].text);
      getusersEmails.push(itemsArr[item].secondaryText);
    }
    itemsArr.forEach((e) => {
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
  };

  const validateSiteOwner = (itemsSiteOwner: any[]) => {
    // show error message if this is a guest user
    if (itemsSiteOwner.length > 0) {
      let userEmail = itemsSiteOwner[0].secondaryText.toLowerCase();
      if (
        userEmail.indexOf("cohnreznick.com") == -1 &&
        userEmail.indexOf("cohnreznickdev") == -1
      ) {
        // this is a guest user, do not validate
        // this.setState({ addusers: [] });
      } else {
        getPeoplePickerItems(itemsSiteOwner);
      }
    } else {
      // this.setState({ addusers: [] });
    }
  };

  // const getMatterNumbersForClientSite = async (clientSiteNumber?: any) => {
  //   const hubSite = Web(GlobalValues.HubSiteURL);
  //   // get the current site's relative URL
  //   // const clientSiteServerRelativeUrl = spContext._pageContext._web.serverRelativeUrl;

  //   let engagementList = await hubSite.lists
  //     .getByTitle("Engagement List")
  //     .items.select(
  //       "Title",
  //       "ClientNumber",
  //       "EngagementName",
  //       "ID",
  //       "WorkYear",
  //       "Team",
  //       "Portals_x0020_Created"
  //     )
  //     .getAll();

  //   // console.table(engagementList);

  //   // TODO: finish filtering the list based on the selected team and client site number
  //   const filterByPortalsCreatedTax = engagementList.filter((listItem) => {
  //     return (
  //       listItem.Portals_x0020_Created === null &&
  //       listItem.Team == "Tax" &&
  //       listItem.ClientNumber == clientSiteNumber
  //     );
  //   });

  //   console.table(filterByPortalsCreatedTax);
  // };

  // runs after a radio button is selected for Team type

  useEffect(() => {
    console.log("team selected::", team);

    // getMatterNumbersForClientSite();
  }, [team]);

  useEffect(() => {
    console.log("portalSelected::", portalSelected);
  }, [portalSelected]);

  useEffect(() => {
    console.log("dateSelections::", dateSelections);
  }, [dateSelections]);

  useLayoutEffect(() => {
    let taxPortalInfo = [];
    let audPortalInfo = [];

    if (isBulkRolloverOpen) {
      console.log(
        "useLayoutEffect fired, calling getMatterNumbersForClientSite::"
      );

      getMatterNumbersForClientSite(clientSiteNumber).then((response) => {
        console.log("logging response from getMatterNumbersForClientSite::", response);

        // setItems(response);
        // setItemsStaged(response);
      });

      // console.log('logging taxAndAudMatters from BulkRollover useLayoutEffect::', taxAndAudMatters);

    }
  }, [isBulkRolloverOpen]);

  // define columns/viewfields so the ListView component knows what to render
  const viewFields: IViewField[] = [
    // {
    //   name: "ID",
    //   displayName: "",
    //   sorting: false,
    //   minWidth: 0,
    //   maxWidth: 0,
    //   isResizable: false,
    // },
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
    // {
    //   name: "SiteOwner",
    //   displayName: "Site Owner",
    //   sorting: false,
    //   minWidth: 100,
    //   maxWidth: 250,
    //   isResizable: true,
    //   render: (rowItem, index, column) => {
    //     return (
    //       <div>
    //         <PeoplePicker
    //           context={spContext}
    //           showtooltip={false}
    //           required={true}
    //           onChange={(items) => validateSiteOwner(items)}
    //           showHiddenInUI={false}
    //           principalTypes={[PrincipalType.User]}
    //           ensureUser={true}
    //           personSelectionLimit={1}
    //           placeholder="Enter name or email"
    //           // defaultSelectedUsers={this.state.addusers}
    //         />
    //       </div>
    //     );
    //   },
    // },
    // {
    //   name: "PortalExpiration",
    //   displayName: "Expiration Date",
    //   sorting: false,
    //   minWidth: 100,
    //   maxWidth: 250,
    //   isResizable: true,
    //   render: (rowItem, index, column) => {
    //     console.log("rowItem::", rowItem);

    //     const onDateChange = (date: Date | null | undefined): void => {
    //       setDateSelections((prevSelections) => ({
    //         ...prevSelections,
    //         [rowItem.ID]: date, // Use a unique identifier from your row data
    //       }));
    //     };

    //     return (
    //       <div>
    //         <DatePicker
    //           // label="DateTime Picker - 24h"
    //           // dateConvention={DateConvention.DateTime}
    //           // timeConvention={TimeConvention.Hours24}
    //           allowTextInput={false}
    //           value={new Date(rowItem.PortalExpiration)}
    //           initialPickerDate={new Date()}
    //           onSelectDate={onDateChange}
    //           formatDate={onFormatDate}
    //         />
    //       </div>
    //     );
    //   },
    // },
  ];

  // define columns/viewfieldsStaged so the ListView component knows what to render (is for staged portals)
  const viewFieldsStaged: IViewField[] = [
    // {
    //   name: "ID",
    //   displayName: "",
    //   sorting: false,
    //   minWidth: 0,
    //   maxWidth: 0,
    //   isResizable: false,
    // },
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
              onChange={(item) => validateSiteOwner(item)}
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
        // console.log("rowItem::", rowItem);

        const onDateChange = (date: Date | null | undefined): void => {
          setDateSelections((prevSelections) => ({
            ...prevSelections,
            [rowItem.ID]: date, // Use a unique identifier from your row data
          }));
        };

        return (
          <div>
            <DatePicker
              // label="DateTime Picker - 24h"
              // dateConvention={DateConvention.DateTime}
              // timeConvention={TimeConvention.Hours24}
              allowTextInput={false}
              value={new Date(rowItem.PortalExpiration)}
              initialPickerDate={new Date()}
              onSelectDate={onDateChange}
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
        minWidth={1200}
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
        <div className={styles.choiceGroupContainer}>
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
              Select engagements below to bulk rollover. No permissions will be
              rolled over to the new portals.
            </span>

            {/* ListView component to hold portals available for rollover */}
            <div className={styles.listViewPortsForRollover}>
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
            </div>


            <span className={styles.guidanceText}>
              Enter a Site Owner and Expiration Date for each portal to
              rollover. No permissions will be rolled over to the new portals.
            </span>

            {/* ListView component to hold staged portals */}
            <ListView
              items={itemsStaged}
              viewFields={viewFieldsStaged}
              // iconFieldName="FileRef"
              compact={true}
              selectionMode={SelectionMode.none}
              // selection={(selectionItem) => setPortalSelected(selectionItem)}
              // defaultSelection={defaultSelectedFromScreen2}
              showFilter={false}
              key="engagementPortalsStaged"
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
