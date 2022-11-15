import * as React from "react";
import {
    Toggle, TextField, Label, CommandButton, Icon
} from 'office-ui-fabric-react';
import { IAssuranceEngSplitRollover } from "./ICreateEngagement";
import styles from "../ClientInfoWebpart.module.scss";
import { initializeIcons } from 'office-ui-fabric-react';
import { AssuranceSplitRolloverModel } from "../../Dataprovider/AssuranceSplitRollover";
import { GlobalValues } from "../../Dataprovider/GlobalValue";

initializeIcons();

class AssuranceEngSplitRollover extends React.Component<IAssuranceEngSplitRollover> {
    public state = {
        Data: []
    };

    public uuidv4() {
        return GlobalValues.uuidv4String.replace(/[xy]/g, (c) => {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    public componentDidMount() {
        let data = [...this.props.Data];
        this.setState({ Data: data });
    }

    public CreateNewRow = () => {
        let obj = new AssuranceSplitRolloverModel();
        let data = this.state.Data;
        obj.IsNewRecord = true;
        obj.CreateasRollover = true;
        obj.IsSplitEngRollOver = false;
        obj.RecordId = this.uuidv4();
        obj.CRETGroupUsers = [];
        obj.CLGroupUsers = [];
        obj.NewSplitValue = "";

        data.push(obj);
        this.setState({
            Data: data
        });
        this.props.Control(this.state.Data);
    }

    public RemoveRow = (e, elementRecordId) => {
        let data = this.state.Data;
        let filteredArray = data.filter(item => item.RecordId == elementRecordId)[0];
        let index = data.indexOf(filteredArray);
        if (index !== -1) {
            data.splice(index, 1);
            this.setState({ Data: data });
        }
        this.props.Control(this.state.Data);
    }

    public NewTextValue = (ev, newValue, element) => {
        let data = this.state.Data;
        data.filter(e => e.RecordId == element)[0].NewSplitValue = newValue;
        this.setState({ Data: data });

        this.props.Control(this.state.Data);


    }

    public ToggleonChange(ev, checked, element) {

        let data = this.state.Data;
        data.filter(e => e.RecordId == element)[0].CreateasRollover = checked;
        this.setState({ Data: data });
        this.props.Control(this.state.Data);
    }


    public onBlurSplitName = (e, index) => {
        const regex = /^[0-9a-zA-Z]+$/;
        const { value, id } = e.target;
        if (value == undefined) {
            return value == "";
        }
        let data = this.state.Data;

        if (!value.match(regex) || value === "") {
            data.filter(en => en.RecordId == index)[0].NewSplitValue = "";
            data.filter(em => em.RecordId == index)[0].Errormessage = "Only Alphanumeric values are allowed";
        }
        else if (data.filter(ev => ev.NewSplitValue.toString().toLowerCase() == value.toString().toLowerCase()).length > 1) {
            data.filter(en => en.RecordId == index)[0].NewSplitValue = "";
            data.filter(em => em.RecordId == index)[0].Errormessage = "Already Exists";
        }
        else {
            data.filter(en => en.RecordId == index)[0].NewSplitValue = value;
            data.filter(em => em.RecordId == index)[0].Errormessage = "";
        }
        this.setState({ Data: data });
        this.props.Control(this.state.Data);
    }

    public render(): React.ReactElement<IAssuranceEngSplitRollover> {

        return (
            <div className={styles.splitRollover}>
                <div className={styles.spiltcolLabels}>
                    <div>
                        <Label>Rollover portal</Label>
                    </div>
                    <div>
                        <Label>Former split names</Label>
                    </div>
                    <div>
                        <Label>New split names</Label>
                    </div>
                </div>
                {this.state.Data.map((element: AssuranceSplitRolloverModel) => {

                    return <div className={styles.splitdivBox}>
                        <Toggle
                            id={element.ToggleId}
                            onText="Rollover"
                            offText="Rollover"
                            hidden={element.IsNewRecord}
                            checked={element.CreateasRollover}
                            onChange={(e, checked) => this.ToggleonChange(e, checked, element.RecordId)}
                        >
                        </Toggle>
                        <TextField disabled={true} value={element.OldSplitValue} hidden={element.IsNewRecord}></TextField>
                        <TextField id={element.TxtId} value={element.NewSplitValue}
                            maxLength={10}
                            errorMessage={element.Errormessage}
                            onBlur={(e) => this.onBlurSplitName(e, element.RecordId)}
                            disabled={!element.CreateasRollover}
                            onChange={(e, newValue) => this.NewTextValue(e, newValue, element.RecordId)}></TextField>

                        {element.IsNewRecord ?

                            <Icon iconName="ChromeClose" onClick={(e) => this.RemoveRow(e, element.RecordId)} className={styles.Icon} />

                            : ""}
                    </div>;


                })
                }

                <div className={styles.newsplitbutton}><Icon iconName="Add" className={styles.Icon} /> <CommandButton text="Add new split" onClick={this.CreateNewRow} /> </div>
            </div >
        );
    }

}
export default AssuranceEngSplitRollover;
