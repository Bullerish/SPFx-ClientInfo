import * as React from "react";
import {
    Checkbox
} from 'office-ui-fabric-react';
import { IAssuranceEngSplitRollover } from "./ICreateEngagement";
import styles from "../ClientInfoWebpart.module.scss";
import { initializeIcons } from 'office-ui-fabric-react';
import { AssuranceSplitRolloverModel } from "../../Dataprovider/AssuranceSplitRollover";

initializeIcons();

class AssuranceEngSplitRolloverUsers extends React.Component<IAssuranceEngSplitRollover> {
    public state = {
        Data: []
    };

    public componentDidMount() {
        let data = [...this.props.Data];
        this.setState({ Data: data });
    }

    public onChangeEmailCRList = (value, email, CRETGroupName) => {

        let CRList = this.state.Data;
        CRList.filter(ev => ev.CreateasRollover == true && ev.CRETGroupName == CRETGroupName).forEach((e) => {

            for (let i = 0; i < e.CRETGroupUsers.length; i++) {
                if (e.CRETGroupUsers[i].email == email && value) {
                    e.CRETGroupUsers[i].checked = true;
                }

                if (e.CRETGroupUsers[i].email == email && !value) {
                    e.CRETGroupUsers[i].checked = false;
                }
            }
        });
        this.setState({ Data: CRList });
        this.props.Control(this.state.Data);
    }

    public onChangeEmailCLList = (value, email, CLGroupName) => {
        let CLList = this.state.Data;
        CLList.filter(ev => ev.CreateasRollover == true && ev.CLGroupName == CLGroupName).forEach((e) => {
            for (let i = 0; i < e.CLGroupUsers.length; i++) {
                if (e.CLGroupUsers[i].email == email && value) {
                    e.CLGroupUsers[i].checked = true;
                }
                if (e.CLGroupUsers[i].email == email && !value) {
                    e.CLGroupUsers[i].checked = false;
                }
            }
        });

        this.setState({ Data: CLList });
        this.props.Control(this.state.Data);

    }

    public render(): React.ReactElement<IAssuranceEngSplitRollover> {
        return (
            <div>
                {this.state.Data.filter(e => e.CreateasRollover == true).map((element: AssuranceSplitRolloverModel) => {

                    return <div className={styles.userLists}>
                        <div className={styles.usergroups}>
                            <span>{element.CRETGroupName}</span>
                            {element.CRETGroupUsers.map(CRETele =>
                                <Checkbox label={CRETele.email} checked={CRETele.checked} onChange={(ev, value) => {
                                    this.onChangeEmailCRList(value, CRETele.email, element.CRETGroupName);
                                }} />
                            )
                            }
                        </div>
                        <div className={styles.usergroups}>
                            <span>{element.CLGroupName}</span>
                            {element.CLGroupUsers.map(CLele =>
                                <Checkbox label={CLele.email} checked={CLele.checked} onChange={(ev, value) => {
                                    this.onChangeEmailCLList(value, CLele.email, element.CLGroupName);
                                }} />
                            )
                            }
                        </div>
                    </div>;
                })
                }
            </div >
        );
    }
}
export default AssuranceEngSplitRolloverUsers;
