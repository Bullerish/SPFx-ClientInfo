import * as React from 'react'
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

import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { IItemAddResult } from "@pnp/sp/items";




const BulkRollover = ({
  spContext,
  isBulkRolloverOpen,
  onBulkRolloverModalHide,
}): React.ReactElement => {
  const [team, setTeam] = useState<string>("");

  useEffect(() => {
    console.log("team selected::", team);
  }, [team]);


const resetState = () => {
  alert("resetState fired::");
  onBulkRolloverModalHide(false);
};

const onTeamChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
  console.log("onTeamChange fired::");
  console.log(option.key);
  setTeam(option.key);
};



  return (
    <>
      <Dialog
        hidden={!isBulkRolloverOpen}
        onDismiss={resetState}
        minWidth={750}
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

export default BulkRollover
