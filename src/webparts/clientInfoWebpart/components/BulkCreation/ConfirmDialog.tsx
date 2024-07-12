import React from 'react';
import { DefaultButton, PrimaryButton, Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react';

interface ConfirmDialogProps {
  hidden: boolean;
  onDismiss: () => void;
  onBack: () => void;
  onCreate: () => void;
  selectedEngagements: any[];
}

const ConfirmDialog: React.FC<ConfirmDialogProps> = ({ hidden, onDismiss, onBack, onCreate, selectedEngagements }) => {
  return (
    <Dialog
      hidden={hidden}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: 'Confirm Bulk Subportal Creation',
        subText: 'Review your selections before creating subportals',
      }}
      modalProps={{
        isBlocking: true,
      }}
    >
      <div>
        {selectedEngagements.map((engagement) => (
          <div key={engagement.ID}>{engagement.Title}</div>
        ))}
      </div>
      <DialogFooter>
        <PrimaryButton text="Create Portals" onClick={onCreate} />
        <DefaultButton text="Back" onClick={onBack} />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};

export default ConfirmDialog;
