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
// import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { sp } from "@pnp/sp";
import { IFieldAddResult } from "@pnp/sp/fields/types";
import "@pnp/sp/site-users";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteUser } from "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import StatusDialog from "./StatusDialog";

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
      const subWebs = await sp.web
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
      console.log("refetch of subwebs occured::::");
      setItems(subWebsWithKey);
    }

    async function getCurrentUserId() {
      // const userId = await sp.web.currentUser();
      const userId = await sp.web.currentUser();
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

              // console.log("subportal type: ", subPortalType);

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

    itemsToBeAddedForAlerts.forEach((el) => {
      itemDetailsToBeSaved.push(el.ServerRelativeUrl);
    });

    alertsToDelete.forEach((el) => {
      if (el.alertId) {
        itemDetailsToBeDeleted.push(el.alertId);
      }
    });

    // formulate object to input as payload below
    listItem = {
      Title: currentUserId ? currentUserId.UserPrincipalName : "",
      UserPrincipalName: currentUserId ? currentUserId.UserPrincipalName : "",
      AbsoluteUrl: absoluteUrl,
      AlertType: alertTypeItem.key,
      AlertFrequency: alertFrequencyItem.key,
      AlertsToAdd: itemDetailsToBeSaved.toString().replace(/,/g, ";"),
      AlertsToDelete: itemDetailsToBeDeleted.toString().replace(/,/g, ";"),
    };

    console.log("item details to be saved: ", listItem);

    let itemResult = await sp.web.lists
      .getByTitle(userAlertsList)
      .items.filter(`Title eq '${currentUserId.UserPrincipalName}'`)();

    if (itemResult.length > 0) {
      listItemId = itemResult[0].Id;

      const updateResult = await sp.web.lists
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
      const itemAddResult: IItemAddResult = await sp.web.lists
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
    const alertsListEnsureResult = await sp.web.lists.ensure(userAlertsList);

    if (alertsListEnsureResult.created) {
      console.log("list was created somewhere!!!!!");

      // since list was newly created, need to add all the relevant columns/fields
      const alertsToAddField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addMultilineText("AlertsToAdd", 6, true, false, false, true);
      const alertsToDeleteField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addMultilineText("AlertsToDelete", 6, true, false, false, true);
      const UserPrincipalNameField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addText("UserPrincipalName", 255);
      const absoluteURLField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addText("AbsoluteUrl", 255);
      const alertTypeField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addText("AlertType", 255);
      const alertFrequencyField: IFieldAddResult = await sp.web.lists
        .getByTitle(userAlertsList)
        .fields.addText("AlertFrequency", 255);

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
    console.log('in resetState func::');

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
    setIsSubmissionSuccessful(null);

  };

  // // TODO: testing functionality
  // const updateItemsToBeAddedForAlerts = async () => {
  //   console.log("in updateItemsToBeAddedForAlerts::::");

  //   console.log("logging activeItemArr::: ", activeItemArr);

  //   // const output: any[] = itemsToBeAddedForAlerts.filter((obj1) => {
  //   //   return !activeItemArr.some((obj2) => {
  //   //     return obj1.key === obj2.key;
  //   //   });
  //   // });

  //   // const output = itemsToBeAddedForAlerts.filter(obj => {
  //   //   return activeItemArr.indexOf(obj) === -1;
  //   // });

  //   // console.log("logging output from updateItemsToBeAddedForAlerts: ", newItems);
  // };

  // EVENT HANDLERS BELOW
  // // TODO: finish working with transfer state from staged alerts to be set
  // const onActiveItemChanged = (
  //   item: ISubWeb[],
  //   index: number,
  //   ev: React.FocusEvent<HTMLElement>
  // ) => {
  //   ev.stopPropagation();
  //   console.log("ON ACTIVE ITEM CHANGED FIRING");
  //   // activeItemArr = [];
  //   // activeItemArr.push(item);
  //   // const selectionGetItems = selectionForAlertsToAdd.getItems();
  //   console.log("logging activeItemArr: ", activeItemArr);

  //   const output: any[] = itemsToBeAddedForAlerts.filter((obj) => {
  //     return activeItemArr.indexOf(obj as any) === -1;
  //   });

  //   // const output: any[] = items.filter((obj1) => {
  //   //   return !itemsToBeAddedForAlerts.some((obj2) => {
  //   //     return obj1.key === obj2.key;
  //   //   });
  //   // });

  //   // console.log("logging output from onActiveItemChanged: ", output);
  //   // setAlertsFactored(output);
  //   // setItemsToBeAddedForAlerts(output);
  //   // updateItemsToBeAddedForAlerts(output);

  //   setItems((prevState) => [...prevState, item as any]);
  //   // selectionForAlertsToAdd.setItems(output, false);
  //   // setItemsToBeAddedForAlerts(output);
  // };

  // function that runs when the user enters text into the Filter text box
  const onChangeFilterText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ) => {
    const output: any[] = subWebInfo.filter((obj) => {
      return itemsToBeAddedForAlerts.indexOf(obj) === -1;
    });

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

  // TODO: once user clicks close or x on statusDialog, we close all dialogs/modals and reset state
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
        onDismiss={() => onAlertModalHide(true)}
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
          onChange={() => onChangeFilterText}
          className={styles.filterControlStyles}
        />
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
            onChange={() => onAlertTypeChange}
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
            onChange={() => onAlertFrequencyChange}
            placeholder="Select an option"
            options={[
              { key: "immediately", text: "Send notification immediately" },
              { key: "dailySummary", text: "Send a daily summary" },
              { key: "weeklySummary", text: "Send a weekly summary" },
            ]}
            styles={{ dropdown: { width: 300 } }}
          />
        </div>
        <DialogFooter>
          <PrimaryButton onClick={factorAlertsToDelete} text="Save Alerts" />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
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
