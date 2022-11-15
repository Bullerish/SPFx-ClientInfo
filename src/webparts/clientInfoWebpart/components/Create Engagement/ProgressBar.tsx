import * as React from 'react';
import styles from "../ClientInfoWebpart.module.scss";

import * as OfficeUI from 'office-ui-fabric-react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

interface IProgressBar {
  isVisible: boolean;
  Message: string;
}

class ProgressBar extends React.Component<IProgressBar>  {

  public render() {
    const DeleteDialoag = {
      type: DialogType.normal,
      title: "",
    };
    const modalProps = {
      isBlocking: true,

    };
    return (
      <React.Fragment>
        <Dialog
          hidden={!this.props.isVisible}
          dialogContentProps={DeleteDialoag}
          modalProps={modalProps}
          containerClassName={'ms-dialogMainOverride ' + styles.ProgressBar}>
          <OfficeUI.Spinner
            label={this.props.Message + "..."}
          />
        </Dialog>
      </React.Fragment>
    );
  }
}

export default ProgressBar;