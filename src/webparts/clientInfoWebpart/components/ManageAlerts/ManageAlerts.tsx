import * as React from "react";
import { useState, useEffect } from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import { DefaultButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList'
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { sp } from "@pnp/sp";

export interface ISubWeb {
  Title: string;
  Id: string;
  ServerRelativeUrl: string;
}

export interface IUserInfo {
  Id: number;
  UserPrincipalName: string;
}

// parent container Manage Alerts component
const ManageAlerts = ({ spContext, isAlertModalOpen, onAlertModalHide }) => {
  const [subWebInfo, setSubWebInfo] = useState<ISubWeb[]>([]);
  const [currentAlertsInfo, setCurrentAlertsInfo] = useState<object[]>([]);
  const [currentUserId, setCurrentUserId] = useState<IUserInfo>();

  const hostUrl = window.location.host;
  const alertsArrayInfo: object[] = [];

  // console.log("hosturl: ", hostUrl);

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

  return (
    <div>
      {/* <Modal
        isOpen={isAlertModalOpen}
        onDismiss={() => onAlertModalHide(true)}
        isBlocking={false}
        // containerClassName={styles.container}
        // dragOptions={this.state.isDraggable ? this._dragOptions : undefined}
      > */}
      <Dialog
        hidden={!isAlertModalOpen}
        onDismiss={() => onAlertModalHide(true)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Alerts Management"
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 700 } },
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={() => onAlertModalHide(true)} text="Save" />
          <DefaultButton onClick={() => onAlertModalHide(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ManageAlerts;
