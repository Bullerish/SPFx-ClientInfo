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
import {Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteUser } from "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import StatusDialog from "./StatusDialog";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";

// for subwebs call
export interface ISubWeb {
  key: string;
  Title: string;
  Id: string;
  ServerRelativeUrl: string;
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
  const [isConfirmationHidden, setIsConfirmationHidden] =
    useState<boolean>(true);
  const [isSubmissionSuccessful, setIsSubmissionSuccessful] =
    useState<boolean>();
  const [statusDialogHidden, setStatusDialogHidden] = useState<boolean>();

  const hostUrl: string = window.location.host;
  const absoluteUrl: string = spContext.pageContext._web.absoluteUrl;

  // const clientPortalWeb = Web(absoluteUrl);

  const userAlertsList = "UserAlertsList";
  const alertsArrayInfo: object[] = [];
  const existingAlerts: ISubWeb[] = [];
  const subWebsWithKey: ISubWeb[] = [];

  let itemDetailsToBeSaved: string[] = [];
  let itemDetailsToBeDeleted: string[] = [];
  // let activeItemArr = [];
  let columns: IColumn[] = [
    {
      key: "column1",
      name: "Sub-Portal Name",
      fieldName: "Title",
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


  // useEffect to get Subwebs
  //
  //
  let alertWeb = Web(absoluteUrl);
  useEffect(() => {

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

      subWebs.forEach((subWebItem) => {
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
          subPortalType === "AUD-FE"
        ) {
          if (determinesPortalType === true) {
            typeOfSubPortal = "Workflow";
          } else {
            typeOfSubPortal = "File Exchange";
          }

          let subWebItemWithKey: any = {
            ...subWebItem,
            key: subWebItem.Id,
            subPortalType: subPortalType,
            typeOfSubPortal: typeOfSubPortal,
            matterNumber: matterNumber,
          };
          subWebsWithKey.push(subWebItemWithKey);
        }
      });

      // console.log(subWebsWithKey);
      setSubWebInfo(subWebsWithKey);
      //console.log("refetch of subwebs occured::::");
      setItems(subWebsWithKey);
    }

    async function getCurrentUserId() {
      const userId = await alertWeb.currentUser();
      setCurrentUserId(userId);
    }

    getCurrentUserId();
    getSubwebs();
  }, []);

  // will run only if subWebInfo is changed/Contains API call to Alerts endpoint (fetches existing alerts)
  useEffect(() => {
    let subPortalTypeName: string = "";
    let subPortalTypeFunc: string = "";
    let subPortalType: string = "";
    // const itemsSelection = selection.getItems();
    // let alertsToSet: string[] = [];

    if (subWebInfo.length > 0 && currentUserId) {
      console.log("In Alerts useEffect");
      // get current alerts set for user
      subWebInfo.forEach((item) => {
        fetch(
          `https://${hostUrl}${item.ServerRelativeUrl}/_api/web/alerts?$filter=UserId eq ${currentUserId.Id}`,
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
                item.ServerRelativeUrl.split("/")[3].split("-")[0];
              subPortalTypeFunc =
                item.ServerRelativeUrl.split("/")[3].split("-")[1];
              subPortalType = subPortalTypeName + "-" + subPortalTypeFunc;

              console.log("existing alert data: ", alert.d.results);

              if (
                (subPortalType === "AUD-WF" || subPortalType === "TAX-WF") &&
                alert.d.results.length === 3
              ) {
                // alertsToSet.push(item.Id);
                // testing creating existing alert array
                existingAlerts.push({
                  key: item.key,
                  Id: item.Id,
                  Title: item.Title,
                  ServerRelativeUrl: item.ServerRelativeUrl,
                  subPortalType: item.subPortalType,
                  typeOfSubPortal: item.typeOfSubPortal,
                  matterNumber: item.matterNumber,
                  alertId: alert.d.results[0].ID,
                });
              } else if (
                subPortalType === "AUD-FE" /* || subPortalType === "ADV-FE"*/ &&
                alert.d.results.length === 1
              ) {
                // alertsToSet.push(item.Id);
                existingAlerts.push({
                  key: item.Id,
                  Id: item.Id,
                  Title: item.Title,
                  ServerRelativeUrl: item.ServerRelativeUrl,
                  subPortalType: item.subPortalType,
                  typeOfSubPortal: item.typeOfSubPortal,
                  matterNumber: item.matterNumber,
                  alertId: alert.d.results[0].ID,
                });
              }

              alertsArrayInfo.push(alert);

              setAlertSelectedSubPortals(existingAlerts); //alertsToSet
              setItemsToBeAddedForAlerts(existingAlerts); //alertsToSet
            }
          })
          .catch((error) => {
            console.log(error);
            throw new Error("There has been an error fetching Alerts Data");
          });
      });
      console.log("logging existingAlerts arr: ", existingAlerts);
    }
  }, [subWebInfo, currentUserId]);

  // for setting the pre-existing alerts on the DetailsList UI for user
  //
  //
  useEffect(() => {
    if (isAlertModalOpen) {
      console.log("in If condition for isAlertModalOpen");
      // console.log('logging items state: ', items);

      const output: any[] = items.filter((obj1) => {
        return !itemsToBeAddedForAlerts.some((obj2) => {
          return obj1.key === obj2.key;
        });
      });

      console.log("logging output: ", output);
      setItems(output);
    }
  }, [isAlertModalOpen]);

  // adds listItem either by updating the record or adding a new one if it doesn't already exist for the user
  const addUserAlertsListItem = async () => {
    let listItem: object = {};
    let listItemId: number;

    console.log("in AddUserAlertslistItem Func");

    // remove existing alerts prior to adding list item so we don't create duplicates
    const output: any[] = itemsToBeAddedForAlerts.filter((obj1) => {
      return !alertSelectedSubPortals.some((obj2) => {
        return obj1.key === obj2.key;
      });
    });

    console.log(
      "itemsToBeAddedForAlerts without pre-existing alerts: ",
      output
    );

    output.forEach((el) => {
      itemDetailsToBeSaved.push(el.ServerRelativeUrl);
    });

    alertsToDelete.forEach((el) => {
      if (el.alertId) {
        itemDetailsToBeDeleted.push(el.alertId + "+" + el.ServerRelativeUrl);
      }
    });

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
    };

    console.log("item details to be saved: ", listItem);

    let hubWeb = Web(GlobalValues.HubSiteURL);
    let itemResult = await hubWeb.lists
      .getByTitle(userAlertsList)
      .items.filter(`Title eq '${currentUserId.LoginName}'`)();

    if (itemResult.length > 0) {
      listItemId = itemResult[0].Id;

      const updateResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .items.getById(listItemId)
        .update(listItem);
      console.log("existing item updated", updateResult);

      if (updateResult.data !== (null || undefined)) {
        setIsSubmissionSuccessful(true);
        setStatusDialogHidden(false);
      } else {
        setIsSubmissionSuccessful(false);
        setStatusDialogHidden(false);
      }
    } else {
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

      console.log("item was newly created", itemAddResult);
    }

    // console.log("itemAddResult: ", itemAddResult);
  };

  // checks for UserAlertsList, if it doesn't exist it gets created then columns will be added
  const ensureAlertsListExists = async () => {
    // console.log(selectionDetails);
    let hubWeb = Web(GlobalValues.HubSiteURL);
    const alertsListEnsureResult = await hubWeb.lists.ensure(userAlertsList);

    if (alertsListEnsureResult.created) {
      console.log("list was created somewhere!!!!!");

      // since list was newly created, need to add all the relevant columns/fields
      const alertsToAddField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addMultilineText("AlertsToAdd", 6, true, false, false, true);
      const alertsToDeleteField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addMultilineText("AlertsToDelete", 6, true, false, false, true);
      const UserPrincipalNameField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("UserPrincipalName", 255);
      const absoluteURLField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("AbsoluteUrl", 255);
      const alertTypeField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("AlertType", 255);
      const alertFrequencyField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("AlertFrequency", 255);
      const timeDayField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("TimeDay", 255);
      const timeTimeField: IFieldAddResult = await hubWeb.lists
        .getByTitle(userAlertsList)
        .fields.addText("TimeTime", 255);

      addUserAlertsListItem();
    } else {
      console.log("list already existed!!!");
      addUserAlertsListItem();
    }
  };

  // logic to process for determining if existing alerts are to be deleted
  const factorAlertsToDelete = (): void => {
    console.log("itemsToBeAdded count: ", itemsToBeAddedForAlerts.length);

    const output: any[] = alertSelectedSubPortals.filter((obj1) => {
      return !itemsToBeAddedForAlerts.some((obj2) => {
        return obj1.key === obj2.key;
      });
    });

    // console.log("logging alertSelectedSubPortals: ", alertSelectedSubPortals);

    console.log("itemsToDelete: ", output);
    setAlertsToDelete(output);
    setIsConfirmationHidden(false);
  };

  // reset state to after subwebs call and default settings
  const resetState = (): void => {
    console.log("in resetState func::");

    setItems(subWebInfo);
    setItemsToBeAddedForAlerts(alertSelectedSubPortals);
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
    setIsSubmissionSuccessful(null);
  };


  // function that runs when the user enters text into the Filter text box
  const onChangeFilterText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ) => {
    const output: any[] = subWebInfo.filter((obj) => {
      return itemsToBeAddedForAlerts.indexOf(obj) === -1;
    });

    console.log(output);

    setItems(
      text
        ? output.filter((i) => i.Title.toLowerCase().indexOf(text) > -1)
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

  return (
    <div>
      <Dialog
        hidden={!isAlertModalOpen}
        onDismiss={onSetStatusDialogHidden}
        minWidth={960}
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
        <div className={styles.addDetailsListContainerStyles}>
          <Text variant="mediumPlus">
            Sub-Portals staged for alert creation:
          </Text>
          <MarqueeSelection selection={selectionForAlertsToAdd}>
            <DetailsList
              items={itemsToBeAddedForAlerts}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.onHover}
              setKey="set"
              // onActiveItemChanged={onActiveItemChanged}
              onShouldVirtualize={() => false}
              selectionMode={SelectionMode.multiple}
              // styles={{ root: { height: "500px" } }}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={1}
              selection={selectionForAlertsToAdd}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              // onItemInvoked={onItemInvoked}
            />
          </MarqueeSelection>
        </div>
        <TextField
          label="Filter by Sub-Portal Name:"
          onChange={onChangeFilterText}
          className={styles.filterControlStyles}
        />
        <Text variant="mediumPlus">
          Select Sub-Portals below to stage for alerts:
        </Text>
        <div className={styles.detailsListContainerStyles}>
          <MarqueeSelection selection={selection}>
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
          </MarqueeSelection>
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
        <DialogFooter>
          <PrimaryButton onClick={factorAlertsToDelete} text="Save Alerts" />
          <DefaultButton onClick={onSetStatusDialogHidden} text="Cancel" />
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
        { itemsToBeAddedForAlerts.length > 0 &&
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
}
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
        <DialogFooter>
          <PrimaryButton onClick={ensureAlertsListExists} text="Confirm" />
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
