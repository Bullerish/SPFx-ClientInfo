import * as React from "react";
import { IAssuranceEngSplitRollover } from "./ICreateEngagement";
import styles from "../ClientInfoWebpart.module.scss";
import { initializeIcons } from 'office-ui-fabric-react';
import { AssuranceSplitRolloverModel } from "../../Dataprovider/AssuranceSplitRollover";

initializeIcons();

class AssuranceEngSplitRolloverDisplayUsers extends React.Component<IAssuranceEngSplitRollover> {
    public state = {
        Data: []
    };


    public componentDidMount() {
        let data = [...this.props.Data];
        this.setState({ Data: data });
    }



    public render(): React.ReactElement<IAssuranceEngSplitRollover> {
        return (
            <div>
                {this.state.Data.filter(e => e.CreateasRollover == true).map((element: AssuranceSplitRolloverModel) => {
                    if (element.CRETGroupUsers.filter(ev => ev.checked == true).length > 0 || element.CLGroupUsers.filter(ev => ev.checked == true).length > 0)
                        return <div className={styles.userLists}>
                            {element.CRETGroupUsers.filter(ev => ev.checked == true).length > 0 ?
                                <div className={styles.usergroupscopy}>
                                    <span>{element.CRETGroupName}</span>
                                    {

                                        element.CRETGroupUsers.filter(ev => ev.checked == true).map(el =>

                                            <div className={styles.usersemails}>{el.email}</div>

                                        )
                                    }
                                </div> : ""}
                            {element.CLGroupUsers.filter(ev => ev.checked == true).length > 0 ?
                                <div className={styles.usergroupscopy}>
                                    <span>{element.CLGroupName}</span>
                                    {
                                        element.CLGroupUsers.filter(ev => ev.checked == true).map(ele =>

                                            <div className={styles.usersemails}>{ele.email}</div>

                                        )
                                    }
                                </div> : ""}
                        </div>;

                })
                }

            </div >
        );
    }

}
export default AssuranceEngSplitRolloverDisplayUsers;
