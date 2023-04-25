import * as React from "react";
import { GlobalValues } from "../Dataprovider/GlobalValue";
import styles from "../components/ClientInfoWebpart.module.scss";
import { ClientInfoClass } from "../Dataprovider/ClientInfoClass";
import { ClientInfoState } from "../Dataprovider/AppState";
import { Text, Link, DefaultButton } from "office-ui-fabric-react";
import CreateEngagement from "../components/Create Engagement/CreateEngagement";
import { ErrorDialog } from "./ErrorDialog";
import ManageAlerts from "../components/ManageAlerts/ManageAlerts";
import ClientProfileInfo from "../components/ClientProfileInfo/ClientProfileInfo";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { xor } from "lodash";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { initializeIcons } from "@uifabric/icons";
import toast, { Toaster } from 'react-hot-toast';

initializeIcons();

export interface IApp {
  spContext: any;
}

class App extends React.Component<IApp> {
  public state = {
    ClientInfoState: new ClientInfoState(),
    isModalOpen: false,
    isAlertModalOpen: false,
    isClientProfileInfoModalOpen: false,
    isAlertIconShown: true,
    isToastShown: false,
    isDontRemind: false
  };

  public onDontRemind = () => {
    this.showHideToast(false);
    this.setState({ isDontRemind: true });
  }

  public reminderToast = <div className={styles.toastReminder}>
    <span>
      Please fill out your Client Profile Information
    </span>
    <div className={styles.linkContainer}>
      <Link onClick={() => this.setState({ isClientProfileInfoModalOpen: true })} className={styles.blueLink}>Profile Form</Link>
      <Link onClick={this.onDontRemind} className={styles.orangeLink}>Don't Remind Me</Link>
      <Link onClick={() => toast.dismiss()} className={styles.orangeLink}>Dismiss</Link>
    </div>
  </div>;

  public componentDidMount() {
    // console.log('in componentDidMount');

    GlobalValues.SetValues(this.props.spContext)
      .then((result) => {
        this.LoadData();
      })
      .catch((error) => {
        console.log("componentDidMount:: error: ", error);
        this.ShowHideErrorDialog(true);
      });

  }

  // event handlers to show hide Client Profile Information
  public onClientProfileInfoModalHide = () => {
    this.showHideClientProfileInfoModal(false);
  }

  // show/hide Client Profile Information Modal
  public showHideClientProfileInfoModal = (isVisible: boolean) => {
    this.setState({ isClientProfileInfoModalOpen: isVisible });
  }

  public onAlertModalHide = () => {
    this.showHideAlertsModal(false);
  }

  // show/hide Manage Alerts Modal
  public showHideAlertsModal = (isVisible: boolean) => {
    this.setState({ isAlertModalOpen: isVisible });
  }

  public OnModalHide = () => {
    this.ShowHideErrorDialog(false);
  }

  // for showing/hiding alert icon
  public showHideAlertIcon = (isVisible: boolean) => {
    this.setState({ isAlertIconShown: isVisible });
  }

  public showHideToast = (isVisible: boolean) => {
    console.log('in showHideToast func, value of isVisible is: ', isVisible);
    if (isVisible) {
      toast(
        this.reminderToast,
        {
          icon: <Icon iconName="AlertSolid" className={styles.alertIcon} />,
          duration: Infinity,
          id: 'reminderToast'
        }
      );
    } else {
      console.log('in else block of showHideToast func, dismissing toast');
      toast.dismiss();
    }
  }

  public ShowHideErrorDialog = (isVisible) => {
    this.setState({ isModalOpen: isVisible });
  }

  public LoadData = () => {
    let obj = new ClientInfoClass();
    console.log('load data');
    obj.GetClientInfo().then((results) => {
      let objState = this.state.ClientInfoState;
      objState.UpdateState(objState, results);
      this.setState({
        ClientInfoState: objState.UpdateState(objState, results),
      });
    });
  }

  public render() {
    var url = window.location.href;
    let IsPermissionPage = false;
    if (url.indexOf(GlobalValues.PermissionPage) > -1) IsPermissionPage = true;
    return (
      <React.Fragment>
        <Toaster position="top-right" containerClassName={styles.toastContainer} toastOptions={{
          className: styles.toastBorderTop
         }} />
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
                {IsPermissionPage == false ? (
                  <div className={styles.flexinncontainer}>
                    <div>
                      {/* TODO: implement and set state for alretsolid item. Make sure to pass down props that factor state */}
                      {this.state.isAlertIconShown ?
                        <Icon iconName="AlertSolid" className={styles.iconError} />
                        : null
                      }
                      <Link
                        href={"#"}
                        onClick={() =>
                          this.setState({ isClientProfileInfoModalOpen: true })
                        }
                      >
                        My Profile Information
                      </Link>
                    </div>
                    <div>
                      <Link
                        href={"#"}
                        onClick={() =>
                          this.setState({ isAlertModalOpen: true })
                        }
                      >
                        Manage Alerts
                      </Link>
                    </div>
                    {GlobalValues.isCRADUser ? (
                      <div>
                        <Link
                          href={
                            this.props.spContext.pageContext.web.absoluteUrl +
                            GlobalValues.PermissionPage
                          }
                        >
                          Manage Portal
                        </Link>
                        {/* <Link onClick={this.UpdatePBIReportData} >Refresh Dashboard Data</Link> */}
                        <CreateEngagement
                          spContext={this.props.spContext}
                        ></CreateEngagement>
                      </div>
                    ) : (
                      ""
                    )}

                    {GlobalValues.isCRETUser ? (
                      <div>
                        <Link
                          href={
                            this.props.spContext.pageContext.web.absoluteUrl +
                            GlobalValues.PermissionPage
                          }
                        >
                          Manage Portal
                        </Link>
                      </div>
                    ) : (
                      ""
                    )}
                  </div>
                ) : null}
                {IsPermissionPage ? (
                  <Link href={this.props.spContext.pageContext.web.absoluteUrl}>
                    Back to Client Dashboard
                  </Link>
                ) : null}
              </div>
            </div>
          </div>
        </div>
        {/* Manage Alerts component */}
        <ManageAlerts spContext={this.props.spContext} isAlertModalOpen={this.state.isAlertModalOpen} onAlertModalHide={this.onAlertModalHide} />
        <ClientProfileInfo spContext={this.props.spContext} isClientProfileInfoModalOpen={this.state.isClientProfileInfoModalOpen} onClientProfileInfoModalHide={this.onClientProfileInfoModalHide} showHideAlertIcon={this.showHideAlertIcon} showHideToast={this.showHideToast} isDontRemind={this.state.isDontRemind} />
        <ErrorDialog
          OnModalHide={this.OnModalHide}
          isModalOpen={this.state.isModalOpen}
        ></ErrorDialog>
      </React.Fragment>
    );
  }
/* UNUSED FUNCTION
  public UpdatePBIReportData = async () => {
    try {
      var ClientNumber = GlobalValues.SiteURL.split("/")[4];
      var IsItemExists: boolean = false;

      //GlobalValues._SetupSP();
      let objToSave = {
        Title: ClientNumber,
        SIteUrl: GlobalValues.SiteURL,
        SiteUrl0: {
          Url: GlobalValues.SiteURL,
          Description: GlobalValues.SiteURL,
        },
        IsClient: "true",
      };

      IsItemExists = await sp.web.lists
        .getByTitle("PBIReportUpdate")
        .items.getAll()
        .then(async (data) => {
          if (
            data.filter(
              (x) => x.Title == ClientNumber && x.IsDatarefreshed == true
            ).length == 0
          ) {
            await sp.web.lists
              .getByTitle("PBIReportUpdate")
              .items.add(objToSave)
              .then(async (item) => {
                GlobalValues.errorTitle = "Success";
                GlobalValues.errorMsg =
                  "Your request to refresh Dashboard data has been submitted successfully!!";
                this.ShowHideErrorDialog(true);
                return true;
              });
          } else {
            GlobalValues.errorTitle = "Success";
            GlobalValues.errorMsg =
              "Your request to refresh Dashboard data has been submitted successfully!!";
            this.ShowHideErrorDialog(true);
          }
          return true;
        });
    } catch (error) {
      console.log("Data load Error: " + error);
    }
  }
  */
}

export default App;
