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
import { PrimaryButton, Stack, Text } from "office-ui-fabric-react";

initializeIcons();

const StatusDialog = ({
  isSubmissionSuccessful,
  statusDialogHidden,
  onSetStatusDialogHidden,
}) => {
  let message: string = isSubmissionSuccessful
    ? "Your User Profile Information has been saved."
    : "Failed to save your updated Client Profile Information. Please try again.";

  return (
    <>
      <Dialog
        hidden={statusDialogHidden}
        onDismiss={() => onSetStatusDialogHidden(true)}
        minWidth={500}
        dialogContentProps={{
          type: DialogType.normal,
          title: isSubmissionSuccessful ? "Success" : "Error",
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
          <Stack.Item align="center">
            <Icon
              iconName={isSubmissionSuccessful ? "Completed" : "ErrorBadge"}
              className={
                isSubmissionSuccessful ? styles.iconSuccess : styles.iconError
              }
            />
          </Stack.Item>
        </Stack>

        <DialogFooter>
          <PrimaryButton
            onClick={() => onSetStatusDialogHidden(true)}
            text="Close"
            className={styles.primaryButton}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default StatusDialog;
