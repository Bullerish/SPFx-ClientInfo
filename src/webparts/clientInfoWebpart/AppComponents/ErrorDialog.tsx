import * as React from "react";
import {
    Dialog, DialogType, DialogFooter
} from "office-ui-fabric-react";
import { GlobalValues } from "../Dataprovider/GlobalValue";
import * as OfficeUI from 'office-ui-fabric-react';

export interface IAdvantageError {
    isModalOpen: boolean;
    OnModalHide: any;

}

export class ErrorDialog extends React.Component<IAdvantageError> {
    public Hide = (): void => {
        this.props.OnModalHide();
    }

    public render() {
        const DeleteDialoag = {
            type: DialogType.normal,
            title: GlobalValues.errorTitle,
        };
        const modalProps = {
            isBlocking: true,
        };
        return (
            <Dialog
                hidden={!this.props.isModalOpen}
                dialogContentProps={DeleteDialoag}
                onDismiss={this.Hide}
                modalProps={modalProps} >
                <div>
                    <p>{GlobalValues.errorMsg}</p>
                </div>

                <DialogFooter>
                <React.Fragment>
                  <OfficeUI.DefaultButton text="Cancel" onClick={this.Hide} />
                </React.Fragment>
              </DialogFooter>
            </Dialog >
        );
    }
}