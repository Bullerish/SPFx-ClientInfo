import * as React from "react";
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
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  CheckboxVisibility,
  SelectionMode,
  IObjectWithKey,
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Text } from "office-ui-fabric-react/lib/Text";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { sp } from "@pnp/sp";
import { IFieldAddResult } from "@pnp/sp/fields/types";
import "@pnp/sp/site-users";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteUser } from "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import StatusDialog from "./StatusDialog";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { setBaseUrl } from "office-ui-fabric-react";
import "@pnp/sp/regional-settings";
// import moment from 'moment-timezone';

var moment = require('moment-timezone');

// for subwebs call
export interface ISubWeb {
  key: string;
  title: string;
  id: string;
  serverRelativeUrl: string;
  subPortalType: string;
  typeOfSubPortal: string;
  matterNumber: string;
  alertId?: string;
}

// for currentUser info
export interface IUserInfo {
  Id: number;
  UserPrincipalName: string;
}

export interface IDetailsListBasicExampleState {
  items: ISubWeb[];
  selectionDetails: {};
}

// parent container Manage Alerts component
const ManageAlerts = ({
  spContext,
  isAlertModalOpen,
  onAlertModalHide,
}): JSX.Element => {
  const [subWebInfo, setSubWebInfo] = useState<ISubWeb[]>([]);
  const [currentUserId, setCurrentUserId] = useState<ISiteUserInfo>();
  const [items, setItems] = useState<ISubWeb[]>([]);
  const [itemsToBeAddedForAlerts, setItemsToBeAddedForAlerts] = useState<
    ISubWeb[]
  >([]);
  const [existingAlertItems, setExistingAlertItems] = useState<
    ISubWeb[]
  >([]);
  // const [selectionDetails, setSelectionDetails] = useState<any>([]);
  const [alertSelectedSubPortals, setAlertSelectedSubPortals] = useState<
    ISubWeb[]
  >([]);
  const [alertsToDelete, setAlertsToDelete] = useState<ISubWeb[]>([]);
  const [alertTypeItem, setAlertTypeItem] = useState<IDropdownOption>({
    key: "allChanges",
    text: "All Changes",
  });
  const [alertFrequencyItem, setAlertFrequencyItem] = useState<IDropdownOption>(
    { key: "immediately", text: "Send notification immediately" }
  );
  const [timeDay, setTimeDay] = useState<IDropdownOption>({
    key: "Sunday",
    text: "Sunday",
  });
  const [timeTime, setTimeTime] = useState<IDropdownOption>({
    key: "0",
    text: "12:00 AM",
  });
  // adding new state for UTC time to submit to list item
  const [utcTime, setUtcTime] = useState<Date>(null);


  const [isConfirmationHidden, setIsConfirmationHidden] =
    useState<boolean>(true);
  const [isSubmissionSuccessful, setIsSubmissionSuccessful] =
    useState<boolean>();
  const [statusDialogHidden, setStatusDialogHidden] = useState<boolean>();
  const [isDataLoaded, setIsDataLoaded] = useState<boolean>(null);
  const [noSubWebs, setNoSubWebs] = useState<boolean>(null);

  const hostUrl: string = window.location.host;
  const absoluteUrl: string = spContext.pageContext._web.absoluteUrl;

  // const clientPortalWeb = Web(absoluteUrl);

  const userAlertsList = "UserAlertsList";
  const alertsArrayInfo: object[] = [];
  // const existingAlerts: ISubWeb[] = [];
  const subWebsWithKey: ISubWeb[] = [];

  let itemDetailsToBeSaved: string[] = [];
  let itemDetailsToBeDeleted: string[] = [];
  // let activeItemArr = [];
  let columns: IColumn[] = [
    {
      key: "column1",
      name: "Engagement Name",
      fieldName: "title",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Matter Number",
      fieldName: "matterNumber",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Portal Type",
      fieldName: "typeOfSubPortal",
      minWidth: 100,
      maxWidth: 210,
      isResizable: true,
    },
  ];

  let columnsAlt: IColumn[] = [
    {
      key: "column1",
      name: "Engagement Name",
      fieldName: "title",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Matter Number",
      fieldName: "matterNumber",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Portal Type",
      fieldName: "typeOfSubPortal",
      minWidth: 100,
      maxWidth: 210,
      isResizable: true,
    },
  ];

  let selection: Selection;
  let selectionForAlertsToAdd: Selection;
  let selectionForExistingAlerts: Selection;

  // useEffect to get Subwebs
  //
  //
  let alertWeb = Web(absoluteUrl);
  // implementing hubWeb to call engagement portal list
  let hubWeb = Web(GlobalValues.HubSiteURL);

  useEffect(() => {
    async function getCurrentUserId() {
      const userId = await alertWeb.currentUser();
      setCurrentUserId(userId);
    }

    getCurrentUserId();
  }, []);

  useEffect(() => {
    setIsDataLoaded(false);
    let subPortalTypeName: string = "";
    let subPortalTypeFunc: string = "";
    let subPortalType: string = "";
    let determinesPortalType: boolean;
    let typeOfSubPortal: string = "";

    let matterNumber: string = "";
    let matterPieceOne: string = "";
    let matterPieceTwo: string = "";
    let matterPieceThree: string = "";

    console.log("In getSubwebs useEffect");
    // get sub-portal information
    async function getSubwebs() {
      const subWebs = await alertWeb
        .getSubwebsFilteredForCurrentUser()
        .select("Title", "ServerRelativeUrl", "Id")
        .orderBy("Title", true)();
      // console.table(subWebs);

      // console.log('logging subWebs:: ', subWebs);

      // fetching only active Engagement Portals from the Engagement Portal List
      const activeEngagementPortals = await hubWeb.lists.getByTitle('Engagement Portal List').items.select('Id', 'Title', 'SiteUrl', 'IsActive', 'PortalId').filter("IsActive eq 1").getAll();
      // console.log('logging activeEngagementPortals:: ', activeEngagementPortals);

      // if the user doesn't have permissions to any subportals and the length is 0, set flags to display message
      if (!subWebs.length) {
        setIsDataLoaded(true);
        setNoSubWebs(true);
      }

      subWebs.forEach((subWebItem) => {
        // console.log('logging subWebItem:: ', subWebItem.Title, subWebItem.ServerRelativeUrl);
        // split on serverRelativeUrl to create SubPortalType
        subPortalTypeName =
          subWebItem.ServerRelativeUrl.split("/")[3].split("-")[0];
        subPortalTypeFunc =
          subWebItem.ServerRelativeUrl.split("/")[3].split("-")[1];
        subPortalType = subPortalTypeName + "-" + subPortalTypeFunc;
        // check if string contains WF, if true then Workflow, if false then File Exchange
        determinesPortalType = subPortalType.indexOf("WF") !== -1;

        matterPieceOne =
          subWebItem.ServerRelativeUrl.split("/")[3].split("-")[2];
        matterPieceTwo =
          subWebItem.ServerRelativeUrl.split("/")[3].split("-")[3];
        matterPieceThree =
          subWebItem.ServerRelativeUrl.split("/")[3].split("-")[4];
        matterNumber =
          matterPieceOne + "-" + matterPieceTwo + "-" + matterPieceThree;

        if (
          subPortalType === "AUD-WF" ||
          subPortalType === "TAX-WF" ||
          subPortalType === "AUD-FE" ||
          subPortalType === "TAX-FE"
        ) {
          if (determinesPortalType === true) {
            typeOfSubPortal = "Workflow";
          } else {
            typeOfSubPortal = "File Exchange";
          }

          if (subWebItem.ServerRelativeUrl.indexOf('-K1-') === -1) {

            let subWebItemWithKey: any = {
              // ...subWebItem,
              id: subWebItem.Id,
              serverRelativeUrl: subWebItem.ServerRelativeUrl,
              title: subWebItem.Title,
              key: subWebItem.Id,
              subPortalType: subPortalType,
              typeOfSubPortal: typeOfSubPortal,
              matterNumber: matterNumber,
            };
            subWebsWithKey.push(subWebItemWithKey);
          }

        }
      });

      // loop filter out subWebs that are not an active Engagement Portal
      const finalSubWebsToSet = subWebsWithKey.filter(element => {
        return activeEngagementPortals.some(item => {
          let splitOnServerRelativeUrl = element.serverRelativeUrl.split('/');
          let idOfPortal = splitOnServerRelativeUrl[splitOnServerRelativeUrl.length - 1];
          // console.log('logging idOfPortal:: ', idOfPortal);

          return idOfPortal === item.PortalId;
        });
      });

      // console.log('logging finalSubWebsToSet:: ', finalSubWebsToSet);

      if (finalSubWebsToSet.length) {
        setSubWebInfo(finalSubWebsToSet);
      } else {
        setIsDataLoaded(true);
        setNoSubWebs(true);
      }

      // setItems(subWebsWithKey);
    }

    if (isAlertModalOpen) {
      getSubwebs();
    }
  }, [isAlertModalOpen]);

  // will run only if subWebInfo is changed/Contains API call to Alerts endpoint (fetches existing alerts)
  useEffect(() => {
    let subPortalTypeName: string = "";
    let subPortalTypeFunc: string = "";
    let subPortalType: string = "";
    let existingAlerts: ISubWeb[] = [];

    // const itemsSelection = selection.getItems();
    // let alertsToSet: string[] = [];

    // console.log('logging subWebInfo:: ', subWebInfo);

    if (subWebInfo.length > 0 && currentUserId) {
      // console.log("In Alerts useEffect");
      // get current alerts set for user
      subWebInfo.forEach((item) => {
        fetch(
          `https://${hostUrl}${item.serverRelativeUrl}/_api/web/alerts?$filter=UserId eq ${currentUserId.Id}`,
          {
            headers: {
              Accept: "application/json;odata=verbose",
            },
          }
        )
          .then((data) => {
            return data.json();
          })
          .then((alert) => {
            if (alert.d.results.length > 0) {
              // grab 3-letter acronym (name) and 2 letter function (func) and combine to form I.e. AUD-WF
              subPortalTypeName =
                item.serverRelativeUrl.split("/")[3].split("-")[0];
              subPortalTypeFunc =
                item.serverRelativeUrl.split("/")[3].split("-")[1];
              subPortalType = subPortalTypeName + "-" + subPortalTypeFunc;

              // console.log("existing alert data: ", alert.d.results);

              if (
                (subPortalType === "AUD-WF" || subPortalType === "TAX-WF") &&
                alert.d.results.length > 2
              ) {
                // alertsToSet.push(item.Id);
                // testing creating existing alert array

                existingAlerts.push({
                  key: item.key,
                  id: item.id,
                  title: item.title,
                  serverRelativeUrl: item.serverRelativeUrl,
                  subPortalType: item.subPortalType,
                  typeOfSubPortal: item.typeOfSubPortal,
                  matterNumber: item.matterNumber,
                  alertId: alert.d.results[0].ID,
                });
              } else if (
                subPortalType === "AUD-FE" /* || subPortalType === "ADV-FE"*/ &&
                alert.d.results.length > 0
              ) {
                // alertsToSet.push(item.Id);
                existingAlerts.push({
                  key: item.id,
                  id: item.id,
                  title: item.title,
                  serverRelativeUrl: item.serverRelativeUrl,
                  subPortalType: item.subPortalType,
                  typeOfSubPortal: item.typeOfSubPortal,
                  matterNumber: item.matterNumber,
                  alertId: alert.d.results[0].ID,
                });
              }

              alertsArrayInfo.push(alert);

              setAlertSelectedSubPortals(existingAlerts); //alertsToSet

            }
          })
          .catch((error) => {
            console.log(error);
            throw new Error("There has been an error fetching Alerts Data");
          });
      });
      // console.log("logging existingAlerts arr: ", existingAlerts);
      console.log('about to set ExistingAlerts state:: ');

      setTimeout(() => {
        setExistingAlertItems(existingAlerts);
        setIsDataLoaded(true);
        setNoSubWebs(false);
      }, 400);



    }


  }, [subWebInfo]);

  // for setting the pre-existing alerts on the DetailsList UI for user
  useEffect(() => {

    console.log('running existingAlertItems useEffect::');

    function setFilteredItems() {
      setTimeout(() => {
        setItems(subWebInfo.filter(item => {
          return !existingAlertItems.some(element => {
            return element.key === item.key;
          });
        }));

      }, 500);
    }

    setFilteredItems();
    // console.log('logging output:: ', output);

  }, [existingAlertItems]);

  // TODO: create new function to take user inputs and factor utc time to submit in list item
  const factorUtcTimeVal = async () => {

    // get user timezone data
    const userTimeZone = spContext.pageContext;
    console.log('logging userTimeZone: ', userTimeZone);

    // user selected day and time of day
    const userDayOfWeek = timeDay.key;
    const userInputTime = timeTime.text;

    console.log('logging userInputTime:: ', userInputTime);

    const dayOfWeekToValue = {
      "Sunday": 0,
      "Monday": 1,
      "Tuesday": 2,
      "Wednesday": 3,
      "Thursday": 4,
      "Friday": 5,
      "Saturday": 6,
    };

    const hourToValue = {
      "12:00 AM": 0,
      "1:00 AM": 1,
      "2:00 AM": 2,
      "3:00 AM": 3,
      "4:00 AM": 4,
      "5:00 AM": 5,
      "6:00 AM": 6,
      "7:00 AM": 7,
      "8:00 AM": 8,
      "9:00 AM": 9,
      "10:00 AM": 10,
      "11:00 AM": 11,
      "12:00 PM": 12,
      "1:00 PM": 13,
      "2:00 PM": 14,
      "3:00 PM": 15,
      "4:00 PM": 16,
      "5:00 PM": 17,
      "6:00 PM": 18,
      "7:00 PM": 19,
      "8:00 PM": 20,
      "9:00 PM": 21,
      "10:00 PM": 22,
      "11:00 PM": 23,
    };

    const numericDayOfWeek = dayOfWeekToValue[userDayOfWeek];
    const [userInputHour, period] = userInputTime.split(' ');
    const numericHour = hourToValue[`${userInputHour} ${period}`];

    console.log('logging numericHour:: ', numericHour);

    // calculate UTC time for the specified day and time
    const today = new Date();
    const currentDayOfWeek = today.getDay();
    const daysToAdd = (numericDayOfWeek + 7 - currentDayOfWeek) % 7;

    const userLocalTime = new Date(today);

    if (alertFrequencyItem.key === 'weeklySummary') {
      userLocalTime.setDate(today.getDate() + daysToAdd);
    }

    userLocalTime.setHours(numericHour);
    userLocalTime.setMinutes(0);

    // console.log('userLocalTime with user selected values: ', userLocalTime.toString());
    // const userTimeToSubmit = await hubWeb.regionalSettings.timeZone.utcToLocalTime(userLocalTime.toISOString());
    // const webTimeZone = await hubWeb.regionalSettings.timeZone();

    // log the hubweb's timezone info
    const { Information } = await hubWeb.regionalSettings.timeZone();
    console.log('logging webs timezone info:', Information);

    // // TODO: testing accounting for timezone bias
    const localTime = userLocalTime.getTime();
    const regionTimeOffset = (Information.Bias + Information.StandardBias + Information.DaylightBias) * 60000;
    console.log('logging regionTimeOffset: ', regionTimeOffset);
    const localTimeOffset = userLocalTime.getTimezoneOffset() * 60000;
    console.log('logging localTimeOffset: ', localTimeOffset);

    // factor regionTimeOffset vs. localTimeOffset.
    if (localTimeOffset > regionTimeOffset) {
      userLocalTime.setTime(localTime + (localTimeOffset - regionTimeOffset));
    } else {
      userLocalTime.setTime(localTime + (regionTimeOffset - localTimeOffset));
    }

    console.log('logging userLocalTime: ', userLocalTime);

    const userSelectedDateTime = moment(userLocalTime).tz(Intl.DateTimeFormat().resolvedOptions().timeZone);
    const timeToSubmit = userSelectedDateTime.format();

    console.log('logging timeToSubmit after moment: ', timeToSubmit);


    return timeToSubmit;

  };

  // adds listItem either by updating the record or adding a new one if it doesn't already exist for the user
  const addUserAlertsListItem = async () => {
    let listItem: object = {};
    let listItemId: number;

    console.log("in AddUserAlertslistItem Func::");

    // remove existing alerts prior to adding list item so we don't create duplicates
    const output: any[] = itemsToBeAddedForAlerts.filter((obj1) => {
      return !alertSelectedSubPortals.some((obj2) => {
        return obj1.key === obj2.key;
      });
    });

    // console.log(
    //   "itemsToBeAddedForAlerts without pre-existing alerts: ",
    //   output
    // );

    output.forEach((el) => {
      itemDetailsToBeSaved.push(el.serverRelativeUrl);
    });

    alertsToDelete.forEach((el) => {
      if (el.alertId) {
        itemDetailsToBeDeleted.push(el.alertId + "+" + el.serverRelativeUrl);
      }
    });

    // TODO: call function to factor utc time and return the date obj to the utcTimeVal
    const utcTimeVal = await factorUtcTimeVal();

    // console.log('Logging utcTimeVal returned:: ', utcTimeVal);

    // formulate object to input as payload below
    listItem = {
      Title: currentUserId ? currentUserId.LoginName : "",
      UserPrincipalName: currentUserId ? currentUserId.LoginName : "",
      AbsoluteUrl: absoluteUrl,
      AlertType: alertTypeItem.key,
      AlertFrequency: alertFrequencyItem.key,
      AlertsToAdd: itemDetailsToBeSaved.toString().replace(/,/g, ";"),
      AlertsToDelete: itemDetailsToBeDeleted.toString().replace(/,/g, ";"),
      TimeDay: timeDay.key,
      TimeTime: timeTime.key,
      // TODO: create new property of UTCTime to submit to list
      UtcTimeValue: utcTimeVal
    };

    console.log("item details to be saved: ", listItem);



    // TODO: uncomment all below after date testing
    const itemAddResult: IItemAddResult = await hubWeb.lists
      .getByTitle(userAlertsList)
      .items.add(listItem);

    if (itemAddResult.data) {
      setIsSubmissionSuccessful(true);
      setStatusDialogHidden(false);
    } else {
      setIsSubmissionSuccessful(false);
      setStatusDialogHidden(false);
    }

    console.log("itemAddResult: ", itemAddResult);
  };

  // logic to process for determining if existing alerts are to be deleted
  const factorAlertsToDelete = (): void => {
    console.log('logging alertsToDelete:: ', alertsToDelete);
    // console.log("itemsToDelete: ", output);
    // setAlertsToDelete(output);
    setIsConfirmationHidden(false);
  };

  // reset state to after subwebs call and default settings
  const resetState = (): void => {
    console.log("in resetState func::");

    setItems([]);
    setItemsToBeAddedForAlerts([]);
    setAlertsToDelete([]);
    setAlertTypeItem({
      key: "allChanges",
      text: "All Changes",
    });
    setAlertFrequencyItem({
      key: "immediately",
      text: "Send notification immediately",
    });
    setTimeDay({ key: "Sunday", text: "Sunday" });
    setTimeTime({ key: "0", text: "12:00 AM" });
    setUtcTime(null);
    setIsDataLoaded(false);
    setIsSubmissionSuccessful(null);
  };

  // function that runs when the user enters text into the Filter text box
  const onChangeFilterText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ) => {

    const output: ISubWeb[] = subWebInfo.filter(item => {
      return !existingAlertItems.some(element => {
        return element.key === item.key;
      });
    });

    console.log(output);

    setItems(
      text
        ? output.filter((i) => i.title.toLowerCase().indexOf(text) > -1)
        : output
    );
  };

  // onChange function fired when user changes selection on Alert Type dropdown
  const onAlertTypeChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    console.log("Alert Type Item: ", item);
    setAlertTypeItem(item);
  };

  // onChange function fired when user changes selection on Alert Frequency
  const onAlertFrequencyChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    console.log("Alert Type Item: ", item);
    setAlertFrequencyItem(item);
  };

  // onChange function fired when user changes selection on Day of Week
  const onTimeDayChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    setTimeDay(item);
  };

  // onChange function fired when user changes selection on Day of Week
  const onTimeTimeChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    setTimeTime(item);
  };

  // function to capture items selected by user and adds them to top DetailsList Component and removes them from bottom DetailsList component
  const transferToMainDetailsList = () => {
    console.log("transferToMainDetailsList fired::");
    const selectionItems = selectionForAlertsToAdd.getSelection();
    const selectionGetItems = selectionForAlertsToAdd.getItems();

    // console.log('loggin selectionGetItems: ', selectionGetItems);
    // console.log('newArrayForAddAlerts', newArrayForAddAlerts);

    setItems((prevState) => [...prevState, ...(selectionItems as any[])]);

    const output: any[] = selectionGetItems.filter((obj) => {
      return selectionItems.indexOf(obj) === -1;
    });

    // console.log('logging output: ', output);

    setItemsToBeAddedForAlerts(output);
  };

  // function to capture items selected by user and adds them to top DetailsList Component and removes them from bottom DetailsList component
  const transferToAddAlertsDetailsList = () => {
    console.log("transferToAddAlertsDetailsList fired::");

    const selectionItems = selection.getSelection();
    const selectionGetItems = selection.getItems();

    setItemsToBeAddedForAlerts((prevState) => [
      ...prevState,
      ...(selectionItems as any[]),
    ]);

    const output: any[] = selectionGetItems.filter((obj) => {
      return selectionItems.indexOf(obj) === -1;
    });

    // console.log('logging output: ', output);

    setItems(output);
  };

  const onHandleExistingAlertsSelection = () => {
    const engagementsToDelete: any[] = selectionForExistingAlerts.getSelection();

    setAlertsToDelete(engagementsToDelete);
  };

  // once user clicks close or x on statusDialog, we close all dialogs/modals and reset state
  const onSetStatusDialogHidden = () => {
    setStatusDialogHidden(true);
    setIsConfirmationHidden(true);
    onAlertModalHide(true);

    resetState();
  };
  // END EVENT HANDLERS

  selection = new Selection({
    onSelectionChanged: () => transferToAddAlertsDetailsList(),
    getKey: (item: any) => item.key,
  });

  selectionForAlertsToAdd = new Selection({
    onSelectionChanged: () => transferToMainDetailsList(),
    getKey: (item: any) => item.key,
  });

  selectionForExistingAlerts = new Selection({
    onSelectionChanged: () => onHandleExistingAlertsSelection(),
    getKey: (item: any) => item.key,
  });

  return (
    <div>
      <Dialog
        hidden={!isAlertModalOpen}
        onDismiss={onSetStatusDialogHidden}
        minWidth={1024}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Alerts Management",
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
          className: styles.manageAlerts,
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        {isDataLoaded && !noSubWebs ?
        <>
          <div className={styles.guidanceText}>
            <span>
              Parent level alerts can only be created for Tax/Assurance File Exchange and Workflow subportals
            </span>
          </div>

          <div className={styles.topContainer}>

          <div className={styles.addDetailsListContainerStyles}>
          <Text variant="mediumPlus">
            Selected parent level alerts:
          </Text>

          {itemsToBeAddedForAlerts.length !== 0 &&

            <DetailsList
              items={itemsToBeAddedForAlerts}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.onHover}
              setKey="set"
              // onDidUpdate={handleNoItemsChecked}
              // onActiveItemChanged={onActiveItemChanged}
              onShouldVirtualize={() => false}
              // selectionMode={!itemsToBeAddedForAlerts.length ? SelectionMode.single: SelectionMode.multiple}
              selectionMode={SelectionMode.multiple}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              selection={selectionForAlertsToAdd}
              selectionPreservedOnEmptyClick={false}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
            />

          }
          </div>
          {/* this is the new existing alerts detailsList component */}
          <div className={styles.addDetailsListContainerStyles}>
          <Text variant="mediumPlus">
            Existing Alerts (select items below to delete when saving):
          </Text>

          {existingAlertItems.length !== 0 &&

            <DetailsList
              items={existingAlertItems}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.onHover}
              setKey="set"
              // onDidUpdate={handleNoItemsChecked}
              // onActiveItemChanged={onActiveItemChanged}
              onShouldVirtualize={() => false}
              // selectionMode={!itemsToBeAddedForAlerts.length ? SelectionMode.single: SelectionMode.multiple}
              selectionMode={SelectionMode.multiple}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              selection={selectionForExistingAlerts}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
              />

          }
          </div>

          </div>
          {/* placeholder for end of fragment */}

        <TextField
          label="Filter by Engagement Name:"
          onChange={onChangeFilterText}
          className={styles.filterControlStyles}
        />
        <Text variant="mediumPlus">
          Select Engagements below to stage for alerts:
        </Text>
        <div className={styles.detailsListContainerStyles}>
        {items.length !== 0 &&

            <DetailsList
              items={items}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.onHover}
              setKey="set"
              onShouldVirtualize={() => false}
              selectionMode={SelectionMode.multiple}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              selection={selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
            />

          }
        </div>
        <div className={styles.alertSettingsContainerStyles}>
          <Dropdown
            label="Alert Type"
            selectedKey={alertTypeItem ? alertTypeItem.key : undefined}
            onChange={onAlertTypeChange}
            placeholder="Select an option"
            options={[
              { key: "allChanges", text: "All Changes" },
              { key: "newItemsAdded", text: "New items are added" },
              {
                key: "existingItemsModified",
                text: "Existing items are modified",
              },
              { key: "itemsDeleted", text: "Items are deleted" },
            ]}
            styles={{ dropdown: { width: 300 } }}
          />
          <Dropdown
            label="Alert Frequency"
            selectedKey={
              alertFrequencyItem ? alertFrequencyItem.key : undefined
            }
            onChange={onAlertFrequencyChange}
            placeholder="Select an option"
            options={[
              { key: "immediately", text: "Send notification immediately" },
              { key: "dailySummary", text: "Send a daily summary" },
              { key: "weeklySummary", text: "Send a weekly summary" },
            ]}
            styles={{ dropdown: { width: 300 } }}
          />
        </div>
        <div className={styles.dayTimeSettingsContainerStyles}>
          <Dropdown
            label="Day of Week"
            disabled={alertFrequencyItem.key === "weeklySummary" ? false : true}
            selectedKey={timeDay ? timeDay.key : undefined}
            onChange={onTimeDayChange}
            placeholder="Select an option"
            options={[
              { key: "Sunday", text: "Sunday" },
              { key: "Monday", text: "Monday" },
              { key: "Tuesday", text: "Tuesday" },
              { key: "Wednesday", text: "Wednesday" },
              { key: "Thursday", text: "Thursday" },
              { key: "Friday", text: "Friday" },
              { key: "Saturday", text: "Saturday" },
            ]}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label="Time"
            selectedKey={timeTime ? timeTime.key : undefined}
            disabled={
              alertFrequencyItem.key === "weeklySummary" ||
              alertFrequencyItem.key === "dailySummary"
                ? false
                : true
            }
            onChange={onTimeTimeChange}
            placeholder="Select an option"
            options={[
              { key: "0", text: "12:00 AM" },
              { key: "1", text: "1:00 AM" },
              { key: "2", text: "2:00 AM" },
              { key: "3", text: "3:00 AM" },
              { key: "4", text: "4:00 AM" },
              { key: "5", text: "5:00 AM" },
              { key: "6", text: "6:00 AM" },
              { key: "7", text: "7:00 AM" },
              { key: "8", text: "8:00 AM" },
              { key: "9", text: "9:00 AM" },
              { key: "10", text: "10:00 AM" },
              { key: "11", text: "11:00 AM" },
              { key: "12", text: "12:00 PM" },
              { key: "13", text: "1:00 PM" },
              { key: "14", text: "2:00 PM" },
              { key: "15", text: "3:00 PM" },
              { key: "16", text: "4:00 PM" },
              { key: "17", text: "5:00 PM" },
              { key: "18", text: "6:00 PM" },
              { key: "19", text: "7:00 PM" },
              { key: "20", text: "8:00 PM" },
              { key: "21", text: "9:00 PM" },
              { key: "22", text: "10:00 PM" },
              { key: "23", text: "11:00 PM" },
            ]}
            styles={{ dropdown: { width: 150 } }}
          />
        </div>
        </>
        : isDataLoaded && noSubWebs ?
          <Text variant="mediumPlus">
            No engagements available for parent level alerts.
          </Text>
        :
        <Spinner size={SpinnerSize.large} label="Loading Portal and Alerts Data..." />

        }

           <DialogFooter>
          {isDataLoaded && !noSubWebs ?
            <>
            <PrimaryButton onClick={factorAlertsToDelete} text="Save" className={styles.rightMargin} />
            <DefaultButton onClick={onSetStatusDialogHidden} text="Cancel" />
            </>
            :
            null
          }
           </DialogFooter>
      </Dialog>

      <Dialog
        hidden={isConfirmationHidden}
        onDismiss={() => setIsConfirmationHidden(true)}
        minWidth={500}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Alerts Summary Confirmation",
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        {itemsToBeAddedForAlerts.length > 0 && (
          <div className={styles.confirmationContainerStyles}>
            <Text variant="large" block nowrap>
              Alerts will be added for:
            </Text>
            {/* confirmation DetailsList for itemsToBeAdded */}
            <DetailsList
              items={itemsToBeAddedForAlerts}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.hidden}
              setKey="set"
              compact={true}
              onShouldVirtualize={() => false}
              selectionMode={SelectionMode.none}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              // selection={selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              // checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
            />
          </div>
        )}
        {alertsToDelete.length > 0 && (
          <div className={styles.confirmationContainerStyles}>
            <Text variant="large" block nowrap>
              Alerts will be deleted for:
            </Text>
            {/* confirmation DetailsList for itemsToBeAdded */}
            <DetailsList
              items={alertsToDelete}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.hidden}
              setKey="set"
              compact={true}
              onShouldVirtualize={() => false}
              selectionMode={SelectionMode.none}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              // selection={selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              // checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
            />
          </div>
        )}
        {!itemsToBeAddedForAlerts.length && !alertsToDelete.length &&
          <Text variant="large" block nowrap>
            No alerts have been selected
          </Text>
        }
        <DialogFooter>
          {itemsToBeAddedForAlerts.length || alertsToDelete.length ?
            <PrimaryButton onClick={addUserAlertsListItem} text="Confirm" />
          :
          null
          }
          <DefaultButton
            onClick={() => setIsConfirmationHidden(true)}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
      <StatusDialog
        isSubmissionSuccessful={isSubmissionSuccessful}
        statusDialogHidden={statusDialogHidden}
        onSetStatusDialogHidden={onSetStatusDialogHidden}
      />
    </div>
  );
};

export default ManageAlerts;
