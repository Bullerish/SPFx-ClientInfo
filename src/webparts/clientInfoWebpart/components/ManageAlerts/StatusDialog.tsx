import * as React from "react";
import { useState, useEffect } from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import { initializeIcons } from '@uifabric/icons';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

initializeIcons();





const StatusDialog = ({ isSubmissionSuccessful, statusDialogHidden }) => {

let message: string = '';


if (isSubmissionSuccessful) {
  message = 'Success';
} else if (!isSubmissionSuccessful) {
  message = 'Submission Failed';
}

  return (
    <>
      <Dialog
        hidden={statusDialogHidden}
        onDismiss={() => statusDialogHidden(true)}
        minWidth={500}
        dialogContentProps={{
          type: DialogType.normal,
          title: message,
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <Icon iconName="Completed" className="ms-IconExample" />
      </Dialog>
    </>
  );
};

export default StatusDialog;
