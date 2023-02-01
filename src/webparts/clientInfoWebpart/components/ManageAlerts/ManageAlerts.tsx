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
  display: 'flex',
  justifyContent: 'space-around',
  padding: '15px 0 15px 0'
});

// for subwebs call
export interface ISubWeb {
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
  const [alertTypeItem, setAlertTypeItem] = useState<IDropdownOption>();
  const [alertFrequencyItem, setAlertFrequencyItem] = useState<IDropdownOption>();

  const hostUrl: string = window.location.host;
  const alertsArrayInfo: object[] = [];

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
      key: "column2",
      name: "Relative Path",
      fieldName: "ServerRelativeUrl",
      minWidth: 100,
      maxWidth: 210,
      isResizable: true,
    },
  ];

  // TODO: Create function to use to set a selection

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
      setSubWebInfo(subWebs);
      setItems(subWebs);
    }

    async function getCurrentUserId() {
      const userId = await sp.web.currentUser();
      setCurrentUserId(userId);
    }

    getCurrentUserId();
    getSubwebs();
  }, []);

  // will run only if subWebInfo is changed
  useEffect(() => {
    if (subWebInfo.length > 0 && currentUserId.Id) {
      console.log("In Alerts useEffect");
      // get current alerts set for user
      // TODO: grab ServerRelativeUrl from getSubwebs(), build below fetch with hostUrl var and ServerRelativeUrl to check if current user has an alert set on sub-portal (additional work to be done to check which list in sub-portal)
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
            // console.log(alert);
            if (alert.d.results.length > 0) {
              alertsArrayInfo.push(alert);
            }
          });
      });
      setCurrentAlertsInfo(alertsArrayInfo);
    }
  }, [subWebInfo, currentUserId]);

  // using to test state updates
  useEffect(() => {
    console.log(subWebInfo);
    console.log(currentAlertsInfo);
    console.log("currentUserId: ", currentUserId);
  }, [subWebInfo, currentAlertsInfo, currentUserId]);

  useEffect(() => {
    console.log('Alert Type Item: ', alertTypeItem);
    console.log('Alert Frequency Item: ', alertFrequencyItem);
  }, [alertTypeItem, alertFrequencyItem]);

  // TODO: Need to get selection details and formulate to write to list
  const getSelectionDetails = () => {
    const selectionCount = selection.getSelection();

    console.log(selectionCount);
  };

  const onAlertTypeChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    // console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    // console.log('Alert Type Item: ', item);
    setAlertTypeItem(item);
  };

  const onAlertFrequencyChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    // console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    // console.log('Alert Type Item: ', item);
    setAlertFrequencyItem(item);
  };

  selection = new Selection({
    onSelectionChanged: () => setSelectionDetails(getSelectionDetails()),
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
              setKey="set"
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
              { key: "existingItemsModified", text: "Existing items are modified" },
              { key: "itemsDeleted", text: "Items are deleted" }
            ]}
            styles={{ dropdown: { width: 300 } }}
          />
          <Dropdown
            label="Alert Frequency"
            selectedKey={alertFrequencyItem ? alertFrequencyItem.key : undefined}
            onChange={onAlertFrequencyChange}
            placeholder="Select an option"
            options={[
              { key: "immediately", text: "Send notification immediately" },
              { key: "dailySummary", text: "Send a daily summary" },
              { key: "weeklySummary", text: "Send a weekly summary" }
            ]}
            styles={{ dropdown: { width: 300 } }}
          />
        </div>
        <DialogFooter>
          <PrimaryButton onClick={() => onAlertModalHide(true)} text="Save" />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ManageAlerts;
