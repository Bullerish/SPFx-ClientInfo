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
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
// import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { sp } from "@pnp/sp";
import { IFieldAddResult } from "@pnp/sp/fields/types";
// import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

const detailsListContainerStyles = mergeStyles({
  height: 700,
  overflowY: "scroll",
});

const alertSettingsContainerStyles = mergeStyles({
  display: "flex",
  justifyContent: "space-around",
  padding: "15px 0 15px 0",
});

const filterControlStyles = mergeStyles({
  margin: '0 30px 20px 0',
  maxWidth: '300px'
});

// for subwebs call
export interface ISubWeb {
  key: string;
  Title: string;
  Id: string;
  ServerRelativeUrl: string;
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
  const [currentAlertsInfo, setCurrentAlertsInfo] = useState<object[]>([]);
  const [currentUserId, setCurrentUserId] = useState<IUserInfo>();
  const [items, setItems] = useState<ISubWeb[]>([]);
  const [selectionDetails, setSelectionDetails] = useState<any>([]);
  const [alertSelectedSubPortals, setAlertSelectedSubPortals] = useState<
    string[]
  >([]);
  const [alertTypeItem, setAlertTypeItem] = useState<IDropdownOption>({key: 'allChanges', text: 'All Changes'});
  const [alertFrequencyItem, setAlertFrequencyItem] =
    useState<IDropdownOption>({key: 'immediately', text: 'Send notification immediately'});

  const hostUrl: string = window.location.host;
  const absoluteUrl: string = spContext.pageContext._web.absoluteUrl;
  // const clientPortalWeb = Web(absoluteUrl);

  const userAlertsList = "UserAlertsList";
  const alertsArrayInfo: object[] = [];
  const subWebsWithKey: ISubWeb[] = [];

  // TODO: Assess and complete the implementation of DetailsList
  let selection: Selection;
  let itemDetailsToBeSaved = [];
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
  ];;

  // // TODO: Update columns to reflect matter number and sub-portal type
  // columns: IColumn[] = [
  //   {
  //     key: "column1",
  //     name: "Sub-Portal Name",
  //     fieldName: "Title",
  //     minWidth: 100,
  //     maxWidth: 200,
  //     isResizable: true,
  //   },
  //   {
  //     key: "column2",
  //     name: "Matter Number",
  //     fieldName: "matterNumber",
  //     minWidth: 100,
  //     maxWidth: 200,
  //     isResizable: true,
  //   },
  //   {
  //     key: "column3",
  //     name: "Portal Type",
  //     fieldName: "typeOfSubPortal",
  //     minWidth: 100,
  //     maxWidth: 210,
  //     isResizable: true,
  //   },
  // ];

  // * useEffect to get Subwebs
  useEffect(() => {
    let subPortalTypeName: string = "";
    let subPortalTypeFunc: string = "";
    let subPortalType: string = "";
    let determinesPortalType: boolean;
    let typeOfSubPortal: string = '';

    let matterNumber: string = '';
    let matterPieceOne: string = '';
    let matterPieceTwo: string = '';
    let matterPieceThree: string = '';

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
        determinesPortalType = subPortalType.indexOf('WF') !== -1;

        matterPieceOne = subWebItem.ServerRelativeUrl.split('/')[3].split('-')[2];
        matterPieceTwo = subWebItem.ServerRelativeUrl.split('/')[3].split('-')[3];
        matterPieceThree = subWebItem.ServerRelativeUrl.split('/')[3].split('-')[4];
        matterNumber = matterPieceOne + '-' + matterPieceTwo + '-' + matterPieceThree;

        if (
          subPortalType === "AUD-WF" ||
          subPortalType === "TAX-WF" ||
          subPortalType === "AUD-FE"
        ) {
          if (determinesPortalType === true) {
            typeOfSubPortal = 'Workflow';
          } else {
            typeOfSubPortal = 'File Exchange';
          }

          let subWebItemWithKey = {
            ...subWebItem,
            key: subWebItem.Id,
            subPortalType: subPortalType,
            typeOfSubPortal: typeOfSubPortal,
            matterNumber: matterNumber
          };
          subWebsWithKey.push(subWebItemWithKey);
        }
      });

      // console.log(subWebsWithKey);
      setSubWebInfo(subWebsWithKey);
      setItems(subWebsWithKey);
    }

    async function getCurrentUserId() {
      const userId = await sp.web.currentUser();
      setCurrentUserId(userId);
    }

    getCurrentUserId();
    getSubwebs();
  }, []);

  // * will run only if subWebInfo is changed/Contains API call to Alerts endpoint (fetches existing alerts)
  useEffect(() => {
    let subPortalTypeName: string = "";
    let subPortalTypeFunc: string = "";
    let subPortalType: string = "";
    let alertsToSet: string[] = [];

    if (subWebInfo.length > 0 && currentUserId.Id) {
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

              console.log("subportal type: ", subPortalType);

              if (
                (subPortalType === "AUD-WF" || subPortalType === "TAX-WF") &&
                alert.d.results.length === 3
              ) {
                // console.log("in set items in alerts call");
                // console.log(item.Id);
                alertsToSet.push(item.Id);
              } else if (
                (subPortalType === "AUD-FE" || subPortalType === "ADV-FE") &&
                alert.d.results.length === 1
              ) {
                // console.log(item.Id);
                alertsToSet.push(item.Id);
              }

              alertsArrayInfo.push(alert);
              setAlertSelectedSubPortals(alertsToSet);
            }
          })
          .catch((error) => {
            console.log(error);
            throw new Error("There has been an error fetching Alerts Data");
          });
      });
      setCurrentAlertsInfo(alertsArrayInfo);
      // setSelectionAlertItems(alertsArrayInfo);
    }
  }, [subWebInfo, currentUserId]);

  // * for setting the pre-existing alerts on the DetailsList UI for user
  useEffect(() => {
    if (isAlertModalOpen) {
      console.log("in If condition for isAlertModalOpen");

      setTimeout(() => {
        alertSelectedSubPortals.forEach((alertItem) => {
          // console.log(alertItem);
          selection.setKeySelected(alertItem, true, false);
        });
      }, 500);
    }
  }, [isAlertModalOpen]);

  // only used for tracking/logging state values
  useEffect(() => {
    console.log("in selectionDetails useEffect");
    console.log(selectionDetails);
    console.log('userIdInfo: ', currentUserId);
  }, [selectionDetails]);

  // TODO: Formulate list item info/object and add list item to list
  const addUserAlertsListItem = async () => {
    let listItem: object = {};
    console.log('in AddUserAlertslistItem Func');

    selectionDetails.forEach(el => {
      // console.log(el);
      itemDetailsToBeSaved.push(el.ServerRelativeUrl);
    });

    // formulate object to input as payload below
    listItem = {
      Title: currentUserId.UserPrincipalName,
      UserPrincipalName: currentUserId.UserPrincipalName,
      AbsoluteUrl: absoluteUrl,
      AlertType: alertTypeItem.key,
      AlertFrequency: alertFrequencyItem.key,
      AlertsToAdd: itemDetailsToBeSaved.toString().replace(/,/g, ';'),
      // TODO: factor logic for items to be deleted and input data here similar to AlertsToAdd
      AlertsToDelete: ''
    };

    console.log('item details to be saved: ', listItem);

    const itemAddResult: IItemAddResult = await sp.web.lists.getByTitle(userAlertsList).items.add(listItem);

    console.log('itemAddResult: ', itemAddResult);

  };

  // * checks for UserAlertsList, if it doesn't exist it gets created then columns will be added
  const ensureAlertsListExists = async () => {
    console.log(selectionDetails);

    const alertsListEnsureResult = await sp.web.lists.ensure(
      userAlertsList
    );

    if (alertsListEnsureResult.created) {
      console.log("list was created somewhere!!!!!");

      // since list was newly created, need to add all the relevant columns/fields
      const alertsToAddField: IFieldAddResult =
        await sp.web.lists
          .getByTitle(userAlertsList)
          .fields.addMultilineText(
            "AlertsToAdd",
            6,
            true,
            false,
            false,
            true
          );
          const alertsToDeleteField: IFieldAddResult =
        await sp.web.lists
          .getByTitle(userAlertsList)
          .fields.addMultilineText(
            "AlertsToDelete",
            6,
            true,
            false,
            false,
            true
          );
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

  // TODO: filter items based on sub-portal name
  const onChangeFilterText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    setItems(text ? subWebInfo.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : subWebInfo);
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

  // const getSelectionDetails = () => {
  //   const selectionItems = selection.getSelection();
  //   // console.log(selectionItems);
  //   setSelectionDetails(selectionItems);
  // };

  // init new selection to get each selected sub-portal
  selection = new Selection({
    onSelectionChanged: () => setSelectionDetails(selection.getSelection()), // getSelectionDetails()
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
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <TextField label="Filter by Sub-Portal Name:" onChange={onChangeFilterText} className={filterControlStyles} />
        <div className={detailsListContainerStyles}>
          <MarqueeSelection selection={selection}>
            <DetailsList
              items={items}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.always}
              setKey="set"
              onShouldVirtualize={() => false}
              // onDidUpdate={() => setSelectedSubPortals()}
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
        <div className={alertSettingsContainerStyles}>
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
        <DialogFooter>
          {/* TODO: change save button to call function that checks for alerts list at client level, if no list then create it and then add the alert item */}
          <PrimaryButton onClick={ensureAlertsListExists} text="Save Alerts" />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ManageAlerts;
