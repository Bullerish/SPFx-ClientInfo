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

const addDetailsListContainerStyles = mergeStyles({
  height: 250,
  overflowY: "scroll",
});

const detailsListContainerStyles = mergeStyles({
  height: 400,
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
  const [itemsToBeAddedForAlerts, setItemsToBeAddedForAlerts] = useState<ISubWeb[]>([]);
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
  const existingAlerts: object[] = [];
  const subWebsWithKey: ISubWeb[] = [];

  let selection: Selection;
  let selectionForAlertsToAdd: Selection;
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
      console.log('refetch of subwebs occured::::');
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
                alertsToSet.push(item.Id);
                // testing creating existing alert array
                existingAlerts.push({ subPortId: item.Id, subPortRelativeUrl: item.ServerRelativeUrl, alertId: alert.d.results[0].ID });

              } else if (
                (subPortalType === "AUD-FE"/* || subPortalType === "ADV-FE"*/) &&
                alert.d.results.length === 1
                ) {
                  alertsToSet.push(item.Id);
                  existingAlerts.push({ subPortId: item.Id, subPortRelativeUrl: item.ServerRelativeUrl, alertId: alert.d.results[0].ID });
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
      console.log('logging existingAlerts arr: ', existingAlerts);
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
          // selection.setKeySelected(alertItem, true, false);
          selection.setKeySelected(alertItem, true, false);
        });
      }, 500);
    }
  }, [isAlertModalOpen]);

  // only used for tracking/logging state values
  useEffect(() => {
    // console.log("in selectionDetails useEffect");
    // console.log(selectionDetails);
    console.log('currentAlertsInfo: ', currentAlertsInfo);
    console.log('alert selected subportals: ', alertSelectedSubPortals);
    // console.log('logging items array: ', items);
  });

  // TODO: Need to add Alerts to delete in the payload to the user list
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

  // checks for UserAlertsList, if it doesn't exist it gets created then columns will be added
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

  // function that runs when the user enters text into the Filter text box
  const onChangeFilterText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {

    const output: any[] = subWebInfo.filter(obj => {
      return itemsToBeAddedForAlerts.indexOf(obj) === -1;
    });

    setItems(text ? output.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : output);
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

  // TODO: function to capture items selected by user and adds them to top DetailsList Component and removes them from bottom DetailsList component
  const transferToMainDetailsList = () => {

    const selectionItems = selectionForAlertsToAdd.getSelection();
    const selectionGetItems = selectionForAlertsToAdd.getItems();

    // console.log('loggin selectionGetItems: ', selectionGetItems);
    // console.log('newArrayForAddAlerts', newArrayForAddAlerts);

    setItems(items => [...items, ...selectionItems as any[]]);

    const output: any[] = selectionGetItems.filter(obj => {
      return selectionItems.indexOf(obj) === -1;
    });

    // console.log('logging output: ', output);

    setItemsToBeAddedForAlerts(output);
  };


  // TODO: function to capture items selected by user and adds them to top DetailsList Component and removes them from bottom DetailsList component
  const transferToAddAlertsDetailsList = () => {
    let newArrayForAddAlerts = [];

    const selectionItems = selection.getSelection();
    const selectionGetItems = selection.getItems();

    // console.log('loggin selectionGetItems: ', selectionGetItems);
    // console.log('newArrayForAddAlerts', newArrayForAddAlerts);

    setItemsToBeAddedForAlerts(itemsToBeAddedForAlerts => [...itemsToBeAddedForAlerts, ...selectionItems as any[]]);

    const output: any[] = selectionGetItems.filter(obj => {
      return selectionItems.indexOf(obj) === -1;
    });

    // console.log('logging output: ', output);

    setItems(output);
  };

   // TODO: selection handler for when user selects an item in the top DetailsList comp
   selectionForAlertsToAdd = new Selection({
    onSelectionChanged: () => transferToMainDetailsList(),
    getKey: (item: any) => item.key,
  });

  // this selection controls what sub-portal items get added to the "Alerts to be Added" DetailsList at top of page
  selection = new Selection({
    onSelectionChanged: () => transferToAddAlertsDetailsList(),
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
        <div className={addDetailsListContainerStyles}>
          <MarqueeSelection selection={selectionForAlertsToAdd}>
            <DetailsList
              items={itemsToBeAddedForAlerts}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.always}
              setKey="set"
              onShouldVirtualize={() => false}
              selectionMode={SelectionMode.single}
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
        <TextField label="Filter by Sub-Portal Name:" onChange={onChangeFilterText} className={filterControlStyles} />
        <div className={detailsListContainerStyles}>
          <MarqueeSelection selection={selection}>
            <DetailsList
              items={items}
              columns={columns}
              checkboxVisibility={CheckboxVisibility.always}
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
          <PrimaryButton onClick={ensureAlertsListExists} text="Save Alerts" />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ManageAlerts;
