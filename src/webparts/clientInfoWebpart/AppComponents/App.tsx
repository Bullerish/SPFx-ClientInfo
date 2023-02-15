import * as React from "react";
import { GlobalValues } from "../Dataprovider/GlobalValue";
import styles from "../components/ClientInfoWebpart.module.scss";
import { ClientInfoClass } from "../Dataprovider/ClientInfoClass";
import { ClientInfoState } from "../Dataprovider/AppState";
import { Text, Link, DefaultButton } from "office-ui-fabric-react";
import CreateEngagement from "../components/Create Engagement/CreateEngagement";
import { ErrorDialog } from "./ErrorDialog";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import { xor } from "lodash";

export interface IApp {
  spContext: any;
}

class App extends React.Component<IApp> {
  public state = {
    ClientInfoState: new ClientInfoState(),
    isModalOpen: false,
    isAlertModalOpen: false
  };

  public componentDidMount() {
    GlobalValues.SetValues(this.props.spContext).then((result) => {
      this.LoadData();
    }).catch(error => {
      console.log("componentDidMount:: error: ", error);
      this.ShowHideErrorDialog(true);
    });
  }

  public onAlertModalHide = () => {
    this.showHideAlertsModal(false);
  }

   // show/hide Manage Alerts Modal
   public showHideAlertsModal = (isVisible) => {
    this.setState({ isAlertModalOpen: isVisible });
  }

  public OnModalHide = () => {
    this.ShowHideErrorDialog(false);
  }

  public ShowHideErrorDialog = (isVisible) => {
    this.setState({ isModalOpen: isVisible });
  }

  public LoadData = () => {
    let obj = new ClientInfoClass();
    obj.GetClientInfo().then((results) => {
      let objState = this.state.ClientInfoState;
      objState.UpdateState(objState, results);
      this.setState({
        ClientInfoState: objState.UpdateState(objState, results),
      });
    });
  }

  public render() {
    var url = (window.location.href);
    let IsPermissionPage = false;
    if (url.indexOf(GlobalValues.PermissionPage) > -1)
      IsPermissionPage = true;
    return (
      <React.Fragment>
        <div className={styles.clientInfoWebpart}>
          <div className={styles.engagementTeam}>
            <div className={styles.clientHeading}>
              <div className={styles.clientInfo}>
                <h1>
                  {this.state.ClientInfoState.ClientInformation.LinkTitle}
                </h1>
                <Text className={styles.engagementinfo}>
                  {this.state.ClientInfoState.ClientInformation.ClientNumber}
                </Text>
              </div>
              <div className={styles.manageSubportal}>
                {IsPermissionPage == false ?
                  <div className={styles.flexinncontainer}>
                    <div>
                      <Link href={"#"} onClick={() => this.setState({ isAlertModalOpen: true })}>Manage Alerts</Link>
                    </div>
                    {GlobalValues.isCRADUser ?
                      <div>
                        <Link href={this.props.spContext.pageContext.web.absoluteUrl + GlobalValues.PermissionPage}>Manage Portal</Link>
                        {/* <Link onClick={this.UpdatePBIReportData} >Refresh Dashboard Data</Link> */}
                        <CreateEngagement spContext={this.props.spContext}></CreateEngagement>
                      </div> : ""}

                    {GlobalValues.isCRETUser ?
                      <div>
                        <Link href={this.props.spContext.pageContext.web.absoluteUrl + GlobalValues.PermissionPage}>Manage Portal</Link>
                      </div> : ""}

                  </div> : null}
                {IsPermissionPage ? <Link href={this.props.spContext.pageContext.web.absoluteUrl}>Back to Client Dashboard</Link>
                  : null}
              </div>
            </div>
          </div>
        </div>        
        <ErrorDialog OnModalHide={this.OnModalHide} isModalOpen={this.state.isModalOpen} ></ErrorDialog>
      </React.Fragment>
    );
  }

  public UpdatePBIReportData = async () => {
    try {
      var ClientNumber = GlobalValues.SiteURL.split('/')[4];
      var IsItemExists: boolean = false;

      GlobalValues._SetupSP();
      let objToSave = {
        "Title": ClientNumber,
        "SIteUrl": GlobalValues.SiteURL,
        "SiteUrl0": {
          "Url": GlobalValues.SiteURL,
          "Description": GlobalValues.SiteURL
        },
        "IsClient": "true",
      };

      IsItemExists = await sp.web.lists.getByTitle("PBIReportUpdate").items.getAll().then(async (data) => {
        if((data.filter(x=> x.Title == ClientNumber && x.IsDatarefreshed == true).length == 0)) {
          await sp.web.lists.getByTitle("PBIReportUpdate").items.add(objToSave).then(async (item) => {
            GlobalValues.errorTitle = "Success";
            GlobalValues.errorMsg = "Your request to refresh Dashboard data has been submitted successfully!!";
            this.ShowHideErrorDialog(true);
            return true;
          });
        }
        else {
          GlobalValues.errorTitle = "Success";
          GlobalValues.errorMsg = "Your request to refresh Dashboard data has been submitted successfully!!";
          this.ShowHideErrorDialog(true);
        }
        return true;
      });
    } catch (error) {
      console.log("Data load Error: " + error);
    }
  }
}

export default App;
