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
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { sp } from "@pnp/sp";

const detailsListContainerStyles = mergeStyles({
  height: 700,
  overflowY: "scroll",
});

const alertSettingsContainerStyles = mergeStyles({
  display: "flex",
  justifyContent: "space-around",
  padding: "15px 0 15px 0",
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
const ManageAlerts = ({ spContext, isAlertModalOpen, onAlertModalHide }) => {
  const [subWebInfo, setSubWebInfo] = useState<ISubWeb[]>([]);
  const [currentAlertsInfo, setCurrentAlertsInfo] = useState<object[]>([]);
  const [currentUserId, setCurrentUserId] = useState<IUserInfo>();
  const [items, setItems] = useState<ISubWeb[]>([]);
  const [selectionDetails, setSelectionDetails] = useState<any>();
  const [alertSelectedSubPortals, setAlertSelectedSubPortals] = useState<
    string[]
  >([]);
  const [alertTypeItem, setAlertTypeItem] = useState<IDropdownOption>();
  const [alertFrequencyItem, setAlertFrequencyItem] =
    useState<IDropdownOption>();

  const hostUrl: string = window.location.host;
  const alertsArrayInfo: object[] = [];
  const subWebsWithKey: ISubWeb[] = [];

  // TODO: Assess and complete the implementation of DetailsList
  let selection: Selection;
  let allItems = [];
  let columns: IColumn[];

  // column settings for data being displayed in DetailsList
  columns = [
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
      name: "Id",
      fieldName: "Id",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Relative Path",
      fieldName: "ServerRelativeUrl",
      minWidth: 100,
      maxWidth: 210,
      isResizable: true,
    },
  ];

  // useEffect to get Subwebs
  useEffect(() => {
    console.log("In getSubwebs useEffect");
    // get sub-portal information
    async function getSubwebs() {
      const subWebs = await sp.web
        .getSubwebsFilteredForCurrentUser()
        .select("Title", "ServerRelativeUrl", "Id")
        .orderBy("Created", false)();
      // console.table(subWebs);

      subWebs.forEach((subWebItem) => {
        let subWebItemWithKey = { ...subWebItem, key: subWebItem.Id };
        subWebsWithKey.push(subWebItemWithKey);
      });

      console.log(subWebsWithKey);
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

  // ^ will run only if subWebInfo is changed/Contains API call to Alerts endpoint
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
                console.log("in set items in alerts call");
                console.log(item.Id);
                alertsToSet.push(item.Id);
              } else if (
                (subPortalType === "AUD-FE" || subPortalType === "ADV-FE") &&
                alert.d.results.length === 1
              ) {
                console.log(item.Id);
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

  // using to test state updates
  useEffect(() => {
    console.log('isAlertModalOpen Value: ', isAlertModalOpen);
    if (isAlertModalOpen) {
      console.log('in If condition for isAlertModalOpen');

      setTimeout(() => {
        alertSelectedSubPortals.forEach((alertItem) => {
          console.log(alertItem);
          selection.setKeySelected(alertItem, true, false);
        });
      }, 1000);
    }

  }, [isAlertModalOpen]);


  // TODO: create function to set subportals from alerts to selected for user
  // const setSelectedSubPortals = () => {
  //   if (isAlertModalOpen) {
  //     console.log('in If condition for isAlertModalOpen');
  //     alertSelectedSubPortals.forEach((alertItem) => {
  //       console.log(alertItem);
  //       selection.setKeySelected(alertItem, true, false);
  //     });
  //   }
  // };

  // onChange function fired when user changes selection on Alert Type dropdown
  const onAlertTypeChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    // console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    // console.log('Alert Type Item: ', item);
    setAlertTypeItem(item);
  };

  // onChange function fired when user changes selection on Alert Frequency
  const onAlertFrequencyChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    // console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    // console.log('Alert Type Item: ', item);
    setAlertFrequencyItem(item);
  };

  // TODO: Need to get selection details and formulate to write to list
  const getSelectionDetails = () => {
    const selectionItems = selection.getSelection();

    console.log(selectionItems);
  };

  selection = new Selection({
    onSelectionChanged: () => setSelectionDetails(getSelectionDetails()),
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
          <PrimaryButton
            onClick={() =>
              selection.setKeySelected(
                "d06aa1a8-4523-4ecd-9660-965dedf7732f",
                true,
                false
              )
            }
            text="Save Alerts"
          />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ManageAlerts;
