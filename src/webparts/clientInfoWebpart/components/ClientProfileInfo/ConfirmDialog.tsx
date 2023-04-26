import * as React from "react";
import { useState, useEffect } from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import { initializeIcons } from "@uifabric/icons";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import styles from "../ClientInfoWebpart.module.scss";
import { PrimaryButton, Stack, Text, DefaultButton } from "office-ui-fabric-react";

initializeIcons();

const ConfirmDialog = ({
  confirmDialogHidden,
  onSetConfirmDialogHidden,
  onConfirmSubmission
}) => {
  let message: string = "Confirm Form Submission?";

  return (
    <>
      <Dialog
        hidden={confirmDialogHidden}
        onDismiss={() => onSetConfirmDialogHidden(true)}
        minWidth={500}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Confirmation",
          showCloseButton: true,
          className: styles.statusDialog,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <Stack>
          <Stack.Item align="center">
            <Text variant="large" className={styles.subText}>
              {message}
            </Text>
          </Stack.Item>
        </Stack>

        <DialogFooter>
          <PrimaryButton
            onClick={onConfirmSubmission}
            text="Confirm"
            className={styles.primaryButton}
          />
          <DefaultButton
            onClick={() => onSetConfirmDialogHidden(true)}
            text="Cancel"
            className={styles.defaultButton}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default ConfirmDialog;
