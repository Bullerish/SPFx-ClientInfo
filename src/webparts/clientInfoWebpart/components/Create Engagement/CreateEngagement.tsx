import * as React from "react";
import {
    Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DirectionalHint, Dropdown, IBasePicker, IBasePickerSuggestionsProps, IChoiceGroupOption, Icon, IDropdownOption, IInputProps, ITag, Label, Link, mergeStyleSets, PrimaryButton, Stack, TagPicker, Text, TextField, TooltipHost, TooltipHostBase
} from 'office-ui-fabric-react';
import { ClientInfoClass } from "../../Dataprovider/ClientInfoClass";
import * as OfficeUI from 'office-ui-fabric-react';
import { ICreateEngagement } from "./ICreateEngagement";
import styles from "../ClientInfoWebpart.module.scss";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import { ICamlQuery } from "@pnp/sp/lists";
import { addDays, addMonths } from "office-ui-fabric-react/lib/utilities/dateMath/DateMath";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import AssuranceEngSplit from "./AssuranceEngSplit";
import AssuranceEngSplitRollover from "./AssuranceEngSplitRollover";
import AssuranceEngSplitRolloverUsers from "./AssuranceEngSplitRolloverUsers";
import AssuranceEngSplitRolloverDisplayUsers from "./AssuranceEngSplitRolloverDisplayUsers";
import { AssuranceSplit } from "../../Dataprovider/AssuranceSplit";
import { AssuranceSplitRollover } from "../../Dataprovider/AssuranceSplitRollover";
import { initializeIcons } from 'office-ui-fabric-react';
import ProgressBar from "./ProgressBar";
import { UserAction } from "../../Dataprovider/ActionEnums";
import { ISiteUser } from "@pnp/sp/site-users";
initializeIcons();

const Teamoptions: IChoiceGroupOption[] = [
    { key: 'Assurance', text: 'Assurance' },
    { key: 'Tax', text: 'Tax' },
    { key: 'Advisory', text: 'Advisory' },
];

const TeamoptionsK1: IChoiceGroupOption[] = [
    { key: 'Assurance', text: 'Assurance', disabled: true },
    { key: 'Tax', text: 'Tax' },
    { key: 'Advisory', text: 'Advisory', disabled: true },
];

const TeamoptionsFileExchange: IChoiceGroupOption[] = [
    { key: 'Assurance', text: 'Assurance' },
    { key: 'Tax', text: 'Tax', disabled: true },
    { key: 'Advisory', text: 'Advisory' },
];

const PortalTypeOptions: IChoiceGroupOption[] = [
    { key: 'Workflow', text: 'Workflow' },
    { key: 'File Exchange', text: 'File Exchange' },
    { key: 'K1', text: 'K1' },

];

const PortalTypeOptionsForTax: IChoiceGroupOption[] = [
    { key: 'Workflow', text: 'Workflow' },
    { key: 'File Exchange', text: 'File Exchange', disabled: true },
    { key: 'K1', text: 'K1' },

];

const PortalChoiceOptions: IChoiceGroupOption[] = [
    { key: 'Rollover', text: 'Rollover' },
    { key: 'Create New', text: 'Create New' },
];

const PortalChoiceOptions1: IChoiceGroupOption[] = [
    { key: 'Rollover', text: 'Rollover', disabled: true },
    { key: 'Create New', text: 'Create New' },
];

let today: Date = new Date(Date.now());
let minDate: Date = addMonths(today, 0);
let maxDate: Date = addMonths(today, 12); // for Tax and Assurance
let advMax: Date = addMonths(today, 36); // for Advisory
let portalExpDate: Date = addMonths(today, 18); // ADDED for site expiration different than file expiration

const contentStyles = mergeStyleSets({
    body: {
        flex: "4 4 auto",
        padding: "10px 20px 20px 20px",
        overflowY: "hidden",
        selectors: {
            p: { margin: "14px 0" },
            "p:first-child": { marginTop: 0 },
            "p:last-child": { marginBottom: 0 },
        },
    },
});

let EngagementNameTags: ITag[];
let EngagementNumberTags: ITag[];
let CRN;
let Engagementdata;
let ExDate = "";
let updatedworkyear = false;
let Isnextyear = false;
class CreateEngagement extends React.Component<ICreateEngagement> {
    public state = {
        AdvantagePortalsItems: [],
        IsDataLoaded: false,
        isOpen: false,
        // Dialog states
        validate: false,
        IsDuplicate: false,
        currentScreen: "screen1",
        // Screen 1 Var
        EngagementName: [],
        EngagementNameSelected: "",
        EngagementNameSelected1: [],
        Rollover: false,
        isRollover: false,
        RolloverURL: "",
        EngagementNumber: [],
        EngagementNumberSelected: "",
        UpdatedEngagementNumberSelected: "",
        EngagementNumberSelected1: [],
        EngID: "",
        PortalTypeURL: "",
        TeamURL: "",
        PortalType: "",
        PortalTypeSelected: "",
        Team: "",
        TeamSelected: "",
        Year: "",
        SiteOwner: "",

        // Screen 2 Var
        addusersID: [],
        addusers: [],
        addusers1: [],
        emailaddress: [],
        PortalChoiceSelected: "",

        //Screen 3 Var
        AdvisoryTemplate: [],
        AdvisoryTemplateSelected: "",
        AdvisoryTemplateSelectedKey: "",
        ServiceType: [],
        ServiceTypeSelected: "",
        ServiceTypeSelectedKey: "",
        IndustryType: [],
        IndustryTypeSelected: "",
        IndustryTypeSelectedKey: "",
        Supplemental: [],
        SupplementalSelected: "",
        SupplementalSelectedKey: "",
        subportaladdusersID: [],
        subportaladdusers: [],
        subportalemailaddress: [],
        CRUserList: [],
        CRUserSelected: "",
        CLUserList: [],
        CLUserSelected: "",
        CLPeoplePicker: [],
        AccessUserList: [],
        FinalAccessUserList: "",
        dialogbuttonname: "Next",
        cancelbuttonname: "Cancel",
        titleText: "",
        emailNotification: false,
        portalExpiration: null,
        fileExpiration: null,
        DateExtend: new Date(),
        K1Date: new Date(),
        Message: "",
        showMessageBar: false,
        MessageBarType: OfficeUI.MessageBarType.error,
        disableBtn: false,
        K1FileName: "",
        success: false,
        PortalsCreated: "",
        PortalsCreatedFinal: "",
        PortalId: "",
        Checkeng: false,
        AsuranceSplitData: {
            disabled: true,
            maxval: 0,
            setSliderValue: 0,
            SelectedCategory: "",
            txtValues: [],
            fieldsArray: [],
            sliderVal: 0,
            isSplitNameExist: false,
            minVal: 0,
            splitToggleValue: false
        },
        AssuranceSplitRollover: [],
        WorkpaperPath: "",
        peoplePickerTitle: "Add users for access to this subportal only:",
        showSpinner: false,
        IsPortalEntryCreated: "",
        PreExistingAlertUsers: [],
        UsersToRollAlerts: []
    };

    /**
     * ResetState
     */
    public ResetState = () => {
        this.setState({
            AdvantagePortalsItems: [],
            IsDataLoaded: false,
            isOpen: false,
            // Dialog states
            validate: false,
            IsDuplicate: false,
            currentScreen: "screen1",
            // Screen 1 Var
            EngagementName: [],
            EngagementNameSelected: "",
            EngagementNameSelected1: [],

            Rollover: false,
            isRollover: false,
            RolloverURL: "",

            EngagementNumber: [],
            EngagementNumberSelected: "",
            UpdatedEngagementNumberSelected: "",
            EngagementNumberSelected1: [],
            EngID: "",

            PortalTypeURL: "",
            TeamURL: "",

            PortalType: "",
            PortalTypeSelected: "",

            Team: "",
            TeamSelected: "",

            Year: "",
            SiteOwner: "",

            // Screen 2 Var
            addusersID: [],
            addusers: [],
            addusers1: [],
            emailaddress: [],

            PortalChoiceSelected: "",
            //Screen 3 Var

            AdvisoryTemplate: [],
            AdvisoryTemplateSelected: "",
            ServiceType: [],
            ServiceTypeSelected: "",
            IndustryType: [],
            IndustryTypeSelected: "",


            Supplemental: [],
            SupplementalSelected: "",

            subportaladdusersID: [],
            subportaladdusers: [],
            subportalemailaddress: [],
            CRUserList: [],
            CRUserSelected: "",
            CLUserList: [],
            CLUserSelected: "",

            CLPeoplePicker: [],

            AccessUserList: [],
            FinalAccessUserList: "",

            dialogbuttonname: "Next",
            cancelbuttonname: "Cancel",
            titleText: "",
            emailNotification: false,
            portalExpiration: null,
            fileExpiration: null,
            DateExtend: new Date(),

            Message: "",
            showMessageBar: false,
            MessageBarType: OfficeUI.MessageBarType.error,
            disableBtn: false,

            K1FileName: "",
            success: false,
            PortalsCreated: "",
            PortalsCreatedFinal: "",
            PortalId: "",
            Checkeng: false,
            WorkpaperPath: "",
            AsuranceSplitData: {

                disabled: true,
                maxval: 0,
                setSliderValue: 0,
                SelectedCategory: "",
                txtValues: [],
                fieldsArray: [],
                sliderVal: 0,
                isSplitNameExist: false,
                minVal: 0,
                splitToggleValue: false

            },
            AssuranceSplitRollover: [],
            peoplePickerTitle: "Add users for access to this subportal only:",
            showSpinner: false,
            IsPortalEntryCreated: "",
            PreExistingAlertUsers: [],
            UsersToRollAlerts: []

        });

        updatedworkyear = false;
        Isnextyear = false;
    }

    public openDialog(e) {
        this.ResetState();
        this.loadAssuranceSupplemental();
        e.preventDefault();

        this.setState({
            currentScreen: "screen1",
            isOpen: true,
            dialogbuttonname: "Next",
            titleText: ""

        });

        let absoluteUrl = GlobalValues.SiteURL;
        let finalabsoluteUrl = absoluteUrl.split("/");
        CRN = finalabsoluteUrl[finalabsoluteUrl.length - 1];
    }

    public ShowHideProgressBar = (isVisible) => {
        this.setState({ showSpinner: isVisible });
    }

    public SetAssuranceSplitData = (AsuranceSplitData) => {
        this.setState({ AsuranceSplitData: AsuranceSplitData }, () => {
            if (this.state.AsuranceSplitData.disabled == false) {
                this.setState({ peoplePickerTitle: "Add users for access to these subportals:" });
            }
            else {
                this.setState({ peoplePickerTitle: "Add users for access to this subportal only:" });
            }
        });
    }

    public SetAssuranceSplitDataRollOver = (Data) => {
        this.setState({ AssuranceSplitRollover: Data });
    }

    public loadEngagements = async (Team) => {

        this.setState({
            EngagementNumber: [],
            EngagementNumberSelected: "",
            EngagementNumberSelected1: [],
            EngagementNameSelected: "",
            Year: ""
        });
        this.state.EngagementNumberSelected1.length = 0;
        this.state.EngagementNumber.length = 0;
        let obj = new ClientInfoClass();
        await obj.GetEngInfo().then(async (results) => {
            Engagementdata = results.filter(e => e.Team == Team);
            await Engagementdata.forEach((element) => {
                this.state.EngagementNumber.push({ key: element.Id.toString(), name: element.Title });
            });
        });
        EngagementNumberTags = this.state.EngagementNumber;
    }


    public loadAdvisoryTemplates = () => {
        let obj = new ClientInfoClass();
        let advisoryTemplatesHolder = [];
        obj.GetAdvisoryTemplates().then((results) => {
            results.forEach((element) => {
                advisoryTemplatesHolder.push({ key: element.Id.toString(), text: element.Title });
            });
        });
        this.setState({ AdvisoryTemplate: advisoryTemplatesHolder });
    }


    public loadServiceTypes = () => {
        let obj = new ClientInfoClass();
        let _ServiceType = [];
        obj.GetServiceTypes().then((results) => {
          console.log('logging getservicetypes before sort: ', results);
            results = results
                .slice(0)
                .sort((a, b) =>
                    (false ? a["TemplateTypeOrder"] < b["TemplateTypeOrder"] : a["TemplateTypeOrder"] > b["TemplateTypeOrder"]) ? 1 : -1);

                console.log('logging getservicetypes after sort: ', results);

            results.forEach((element) => {
                if ((element.Title).toLowerCase() == (this.state.TeamSelected).toLowerCase()) {
                    _ServiceType.push({ key: element.Id.toString(), text: element.ServiceType });
                }
            });
            this.setState({ ServiceType: _ServiceType });
        });
        return this.state.ServiceType;
    }

    public loadIndustryTypes = () => {
        let _IndustryType = [];
        let obj = new ClientInfoClass();
        obj.GetIndustryTypes().then((results) => {
            // results = results
            //     .slice(0)
            //     .sort((a, b) =>
            //         (false ? a["WorkYear"] < b["WorkYear"] : a["IndustryType"] > b["IndustryType"]) ? 1 : -1);
            if (this.state.TeamSelected == "Assurance") {
                _IndustryType.push({ key: "N/A", text: "N/A" });
            }
            results.forEach((element) => {
                if (element.Title == this.state.TeamSelected) {
                    _IndustryType.push({ key: element.Id.toString(), text: element.IndustryType });
                }
            });
            this.setState({ IndustryType: _IndustryType });
        });
        return this.state.IndustryType;
    }

    public loadAssuranceSupplemental = () => {
        let obj = new ClientInfoClass();
        const newArray = [];

        obj.GetSupplemental().then((results) => {

            results = results
                .slice(0)
                .sort((a, b) =>
                    (false ? a["Title"] < b["Title"] : a["ID"] > b["ID"]) ? 1 : -1);
            this.state.Supplemental.push({ key: "N/A", text: "N/A" });

            results.forEach(e => {
                if (!newArray.some(o => o.Title === e.Title)) {
                    newArray.push({ ...e });
                }
            });
            newArray.forEach((element) => {
                this.state.Supplemental.push({ key: element.Id.toString(), text: element.Title });
            });
        });
        this.setState({ SupplementalSelected: "N/A", SupplementalSelectedKey: "N/A" });
    }

    public checkEngagement = async (portalsCreated) => {
        if (portalsCreated != null) {
            let finalPortalTypeValue = portalsCreated.split(",");
            let engagementExists = false;
            for (var i = 0; i < finalPortalTypeValue.length; i++) {
                if (finalPortalTypeValue[i] == this.state.PortalTypeURL) {
                    engagementExists = true;
                }
            }
            if (engagementExists == false) {
                this.setState({ PortalsCreatedFinal: portalsCreated + "," + this.state.PortalTypeURL, Checkeng: true });
                return true;
            }
            else {
                if (updatedworkyear == true) {
                    this.setState({ Checkeng: true });
                    return true;
                } else {
                    let ErrorMessage = "You can not create same engagement number " + this.state.EngagementNumberSelected + " for " + this.state.PortalTypeSelected;
                    this.setState({ Message: ErrorMessage, showMessageBar: true, MessageBarType: OfficeUI.MessageBarType.error, Checkeng: false });
                    return false;
                }
            }
        }
        else if (portalsCreated == null) {
            this.setState({ PortalsCreatedFinal: this.state.PortalTypeURL, Checkeng: true });
            return true;
        }
        else {
            let ErrorMessage = "You can not create same engagement number " + this.state.EngagementNumberSelected + " for " + this.state.PortalTypeSelected;
            this.setState({ Message: ErrorMessage, showMessageBar: true, MessageBarType: OfficeUI.MessageBarType.error, Checkeng: false });
            return false;
        }
    }

    private SaveEngagementList = () => {
        let obj = new ClientInfoClass();
        obj.SaveEngagementList(this.state.PortalsCreatedFinal, this.state.EngID).catch((error) => {
            console.log("SaveEngagementList:: Error: ", error);
        });
    }

    private CheckIfEngCreated = (async () => {
        let obj = new ClientInfoClass();
        let FinalEngNumber = updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected;

        return await obj.CheckIfEngCreated(FinalEngNumber).then((results) => {
            this.ShowHideProgressBar(false);
            return results;
        }).catch(error => {
            console.log("CheckIfEngCreated::error:", error);
        });
    });

    public _onChangeEngagementNumber = async (tagList: { key: string, name: string }[]) => {
        if (tagList.length == 0) {
            this.closeMessageBar();
            this.setState({
                EngagementNumberSelected: "",
                EngagementNameSelected: "",
                Year: ""
            });
            this.state.EngagementNumberSelected1.length = 0;

        } else {
            tagList.filter(item => {
                this.setState({
                    EngagementNumberSelected: item.name
                });
                let EngagementNumberEndZero = item.name.slice(-2);
                if (EngagementNumberEndZero == "00") {
                    updatedworkyear = true;
                }
                else {
                    updatedworkyear = false;
                }
                Engagementdata.filter(async (e) => {
                    if (e.Title == item.name) {
                        ExDate = (6) + '-' + (1) + '-' + (parseInt(e.WorkYear) + 2);
                        let dt = new Date(ExDate);
                        const ExDate1: Date = dt;
                        this.setState({
                            /*
                            DateExtend: maxDate,
                            fileExpiration: maxDate,
                            portalExpiration: maxDate,*/
                            K1Date: ExDate1,
                            EngagementNameSelected: e.EngagementName,
                            Year: e.WorkYear,
                            EngID: e.Id,
                            PortalsCreated: e.Portals_x0020_Created,
                            PortalId: e.PortalId
                        });
                        this.checkEngagement(e.Portals_x0020_Created);
                    }
                });
            });
            //this.checkEngagement();
        }
    }

    private newEngagementNumber() {
        if (updatedworkyear == true) {
            let updatedYear = this.state.Year.toString().slice(-2);
            let engNumber = this.state.EngagementNumberSelected.substring(0, this.state.EngagementNumberSelected.length - 2);
            let newEngagementNumber = engNumber + updatedYear;
            this.setState({ UpdatedEngagementNumberSelected: newEngagementNumber });
        }
    }

    public insertdata(siteAbsoluteUrl: string, listname: string, requestdata, requestDigest): Promise<number> {
        console.log('in insertdata func:::');
        let url = `${siteAbsoluteUrl}/_api/web/lists/getbytitle('${listname}')/items`;
        const currWeb = Web(siteAbsoluteUrl);
        return new Promise<number>((resolve, reject) => {
            try {
                fetch(url,
                    {
                        method: "POST",
                        credentials: 'same-origin',
                        headers: {
                            'Accept': 'application/json',
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": requestDigest
                        },
                        body: requestdata,
                    }).then((response) => {
                        return response.json();
                    }).then((response) => {
                        resolve(response.ID);
                    }).catch((error) => {
                        reject(error);
                    });
            }
            catch (e) {
                console.log("insertdata::error", e);
                reject(e);
            }
        });
    }

    public getListItemEntityTypeName(siteAbsoluteUrl: string, listname: string): Promise<string> {
        let url = `${siteAbsoluteUrl}/_api/web/lists/getbytitle('${listname}')?$select=ListItemEntityTypeFullName`;
        return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json;odata=nometadata' } }).then((response) => {
            return response.json();
        }, (errorFail) => {
            console.log("getListItemEntityTypeName::error", errorFail);
        }).then((responseJSON) => {
            return responseJSON.ListItemEntityTypeFullName;
        }).catch((response: any) => {
            console.log("getListItemEntityTypeName::error", response);
            return null;
        });
    }

    public getValues(siteurl: string): any {
        try {
            let url = siteurl + "/_api/contextinfo";
            return fetch(url, {
                method: "POST",
                credentials: 'same-origin',
                headers: { Accept: "application/json;odata=verbose" }
            }).then((response) => {
                let datum = response.json();
                return datum;
            });
        } catch (error) {
            console.log("getValues: " + error);
        }
    }

    public SaveItem(SPUrl, listname) {
        let site = "";
        let PortalId = "";
        let FinalEngNumber = updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected;
        let FinalEngNumberEndZero = updatedworkyear == true ? this.state.EngagementNumberSelected : "";

        // combining the rollover users and new users for rollover portals
        let finalCRUsers = this.state.FinalAccessUserList + this.state.CRUserSelected;

        if (this.state.PortalTypeSelected == "K1") {
            site = GlobalValues.SiteURL + "/TAX-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            PortalId = "TAX-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            return new Promise<number>((resolve, reject) => {
                this.getListItemEntityTypeName(SPUrl, listname)
                    .then((listEntityName) => {
                        const PortalData: any = {
                            '__metadata': {
                                'type': listEntityName
                            },
                            'EngagementName': this.state.EngagementNameSelected,
                            'Title': FinalEngNumber,
                            'EngagementNumberEndZero': FinalEngNumberEndZero,
                            'ClientNumber': CRN,
                            'PortalType': this.state.PortalTypeSelected,
                            'Team': "Tax",
                            'WorkYear': this.state.Year.toString(),
                            'SiteOwnerId': this.state.addusersID,
                            'SiteUrl': { Url: site },
                            'isNotificationEmail': this.state.emailNotification,
                            'PortalExpiration': this.state.K1Date,
                            'PortalId': PortalId
                        };
                        this.getValues(SPUrl)
                            .then((requestDigest) => {
                                this.insertdata(SPUrl, listname, JSON.stringify(PortalData), requestDigest.d.GetContextWebInformation.FormDigestValue)
                                    .then((response) => {
                                        this.CheckIfEngCreated().then((engcrt) => {
                                            if ((response !== null) && (engcrt !== null)) {

                                                this.SaveEngagementList();
                                                resolve(response);
                                                let files = (document.querySelector("#newfile") as HTMLInputElement).files[0];
                                                this.UploadFile(files, FinalEngNumber);
                                                this.setState({ IsPortalEntryCreated: "Y" });
                                            }
                                            else {
                                                this.setState({ IsPortalEntryCreated: "N" });
                                                reject();
                                            }
                                        });
                                    });
                            });
                    });
            });

        }
        else if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Advisory") {
            site = GlobalValues.SiteURL + "/" + this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            PortalId = this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            return new Promise<number>((resolve, reject) => {
                this.getListItemEntityTypeName(SPUrl, listname)
                    .then((listEntityName) => {
                        const PortalData: any = {
                            '__metadata': {
                                'type': listEntityName
                            },
                            'EngagementName': this.state.EngagementNameSelected,
                            'Title': FinalEngNumber,
                            'EngagementNumberEndZero': FinalEngNumberEndZero,
                            'ClientNumber': CRN,
                            'PortalType': this.state.PortalTypeSelected,
                            'Team': this.state.TeamSelected,
                            'WorkYear': this.state.Year.toString(),
                            'SiteOwnerId': this.state.addusersID,
                            'SiteUrl': { Url: site },
                            'ClientMembers': this.state.CLUserSelected,
                            'TemplateType': this.state.AdvisoryTemplateSelected,
                            'isNotificationEmail': this.state.emailNotification,
                            'PortalExpiration': (this.state.portalExpiration ? this.state.portalExpiration : advMax),
                            'PortalId': PortalId
                        };
                        this.getValues(SPUrl)
                            .then((requestDigest) => {
                                this.insertdata(SPUrl, listname, JSON.stringify(PortalData), requestDigest.d.GetContextWebInformation.FormDigestValue)
                                    .then((response) => {
                                        this.CheckIfEngCreated().then((engcrt) => {

                                            if ((response !== null) && (engcrt !== null)) {
                                                this.SaveEngagementList();
                                                resolve(response);
                                                this.setState({ IsPortalEntryCreated: "Y" });
                                            }
                                            else {
                                                this.setState({ IsPortalEntryCreated: "N" });
                                                reject();
                                            }
                                        });
                                    });
                            });
                    });
            });
        }
        else {
            site = GlobalValues.SiteURL + "/" + this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            PortalId = this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
            let RolloverUrl = "";
            let PortalRollOver = false;
            let usersToRollAlerts = '';
            let usersToRollAlertsArray = [];

            if (this.state.PortalChoiceSelected == 'Rollover') {
              this.state.UsersToRollAlerts.forEach(e => {
                usersToRollAlertsArray.push(e.email);
              });

              PortalRollOver = true;
              RolloverUrl = GlobalValues.SiteURL + "/" + this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + this.state.RolloverURL;
              usersToRollAlerts = usersToRollAlertsArray.toString().replace(/,/g, ';');
            }

            // ensuring default expiration dates are set:
            let defaultPortalExpDate = maxDate; // 12 months
            if (this.state.TeamSelected == "Advisory") { defaultPortalExpDate = advMax;} // 36 months
            if (this.state.TeamSelected != "Advisory" && this.state.PortalTypeSelected == "Workflow") { defaultPortalExpDate = portalExpDate;} // 18 months
            let defaultFileExpDate = null;
            if (this.state.TeamSelected != "Advisory" && this.state.PortalTypeSelected == "Workflow") { defaultFileExpDate = maxDate;} // 12 months

            return new Promise<number>((resolve, reject) => {
                this.getListItemEntityTypeName(SPUrl, listname)
                    .then((listEntityName) => {
                        let PortalData: any = {
                            __metadata: {
                                type: listEntityName
                            },
                            'EngagementName': this.state.EngagementNameSelected,
                            'Title': FinalEngNumber,
                            'EngagementNumberEndZero': FinalEngNumberEndZero,
                            'ClientNumber': CRN,
                            'PortalType': this.state.PortalTypeSelected,
                            'Team': this.state.TeamSelected,
                            'WorkYear': this.state.Year.toString(),
                            'SiteOwnerId': this.state.addusersID,
                            'SiteUrl': { Url: site },
                            // Engagement Members will be a combination of rollover users (if rollover and new users)
                            'EngagementMembers': finalCRUsers,// was: this.state.CRUserSelected,
                            // Client Members will only happen in rollover since non CR users are not add-able during portal creation
                            'ClientMembers': this.state.CLUserSelected, // was: this.state.PortalChoiceSelected == "Create New" ? this.state.FinalAccessUserList : this.state.CLUserSelected,
                            'Rollover': PortalRollOver,
                            'RolloverUrl': { Url: RolloverUrl },
                            'IndustryType': this.state.IndustryTypeSelected,
                            'ServiceType': this.state.ServiceTypeSelected,
                            'Supplemental': this.state.SupplementalSelected,
                            'TemplateType': this.state.TeamSelected == 'Tax' ? this.state.ServiceTypeSelected : this.state.AdvisoryTemplateSelected,
                            'isNotificationEmail': this.state.emailNotification,
                            'FileExpiration': this.state.fileExpiration ? this.state.fileExpiration : defaultFileExpDate,
                            'PortalExpiration': this.state.portalExpiration ? this.state.portalExpiration : defaultPortalExpDate,
                            'PortalId': PortalId,
                            'WorkpaperPath': this.state.WorkpaperPath,
                            'UsersToRollAlerts': usersToRollAlerts
                        };
                        console.log('portalData', PortalData);
                        this.getValues(SPUrl)
                            .then((requestDigest) => {
                                if (this.state.TeamSelected == "Assurance" && this.state.PortalTypeSelected == "Workflow" && this.state.AsuranceSplitData.disabled == false) {
                                    let SplitObj = new AssuranceSplit();
                                    this.ShowHideProgressBar(true);
                                    SplitObj.SaveSplitEngagement(PortalData, this.state.AsuranceSplitData).then(value => {
                                        this.CheckIfEngCreated().then((engcrt) => {
                                            if (value == true && engcrt !== null) {
                                                this.SaveEngagementList();
                                                this.ShowHideProgressBar(false);
                                                resolve(1);
                                                this.setState({ IsPortalEntryCreated: "Y" });
                                            }
                                            else {
                                                this.setState({ IsPortalEntryCreated: "N" });
                                                reject();
                                            }
                                        });

                                    }).catch(() => {
                                        this.ShowHideProgressBar(false);
                                    });
                                } else if (this.state.TeamSelected == "Assurance" && this.state.PortalTypeSelected == "Workflow" && this.state.AssuranceSplitRollover.length != 0 && this.state.AssuranceSplitRollover[0].NewSplitValue != null && this.state.PortalChoiceSelected == 'Rollover') {
                                  console.log('in assurance, workflow, and rollover if:::');
                                    let SplitRolloverObj = new AssuranceSplitRollover();
                                    this.ShowHideProgressBar(true);
                                    SplitRolloverObj.SaveSplitEngagementRollover(PortalData, this.state.AssuranceSplitRollover).then(val => {
                                      console.log('in assurance SplitRolloverObj:::');
                                        this.CheckIfEngCreated().then((engcrt) => {

                                            if (val == true && engcrt !== null) {
                                                this.SaveEngagementList();
                                                this.ShowHideProgressBar(false);
                                                resolve(1);
                                                this.setState({ IsPortalEntryCreated: "Y" });
                                            }
                                            else {
                                                reject();
                                                this.setState({ IsPortalEntryCreated: "N" });
                                            }
                                        });
                                    }).catch(() => {
                                        this.ShowHideProgressBar(false);
                                    });

                                }
                                else {
                                  console.log('in else and about to invoke insertdata func:::');
                                    this.insertdata(SPUrl, listname, JSON.stringify(PortalData), requestDigest.d.GetContextWebInformation.FormDigestValue)
                                        .then((response) => {
                                            this.CheckIfEngCreated().then((engcrt) => {
                                                if ((response !== null) && (engcrt !== null)) {
                                                    this.SaveEngagementList();
                                                    resolve(response);
                                                    this.setState({ IsPortalEntryCreated: "Y" });
                                                }
                                                else {
                                                    this.setState({ IsPortalEntryCreated: "N" });
                                                    reject();
                                                }
                                            });
                                        }).catch(() => {
                                            this.ShowHideProgressBar(false);
                                        });
                                }
                            });
                    });
            });
        }
    }

    public UploadFile = (async (file, FinalEngNumber) => {
        let filePrefix = "K1-" + FinalEngNumber + "-";
        let hubWeb = Web(GlobalValues.HubSiteURL);
        await hubWeb.getFolderByServerRelativeUrl(GlobalValues.K1InvestorDocumentsURL).files.add(filePrefix + file.name, file, true).then(async (results) => {
            return await results.file.getItem().then(async (listItem) => {
                listItem.update({
                    EngagementNumber3: FinalEngNumber
                }).then(r => {
                    console.log(file.name + " properties updated successfully!");
                    return true;
                });
            });
        });
    });

    public _onChangeServiceType = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ ServiceTypeSelected: item.text, ServiceTypeSelectedKey: item.key });
    }

    public _onChangeIndustryType = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ IndustryTypeSelected: item.text, IndustryTypeSelectedKey: item.key });
    }

    public _onChangeSupplemental = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ SupplementalSelected: item.text, SupplementalSelectedKey: item.key });
    }

    public _onChangeAdvisoryTemplate = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ AdvisoryTemplateSelected: item.text, AdvisoryTemplateSelectedKey: item.key });
    }
    public _onChangeTeam = async (event: React.FormEvent<HTMLDivElement>, option: IChoiceGroupOption) => {
        this.setState({ TeamSelected: option.text });
        let TeamValue = "";
        if (option.text == "Assurance") {
            TeamValue = "AUD";
            this.setState({ TeamURL: "AUD" });
        }
        if (option.text == "Tax") {
            TeamValue = "TAX";
            this.setState({ TeamURL: "TAX" });
        }
        if (option.text == "Advisory") {
            TeamValue = "ADV";
            this.setState({ TeamURL: "ADV" });
        }
        await this.loadEngagements(TeamValue);

    }
    public _onChangePortalType = async (event: React.FormEvent<HTMLDivElement>, option: IChoiceGroupOption) => {
        this.setState({ PortalTypeSelected: option.text });
        if (option.text == "Workflow") {
            this.setState({ PortalTypeURL: "WF" });
        }
        if (option.text == "File Exchange") {
            this.setState({ PortalTypeURL: "FE" });
            if (this.state.TeamSelected == "Tax") {
                await this.setState({ TeamSelected: "", TeamURL: "" });
            }
        }
        if (option.text == "K1") {
            this.setState({ PortalTypeURL: "K1", TeamSelected: "Tax" });
            this.loadEngagements("TAX");
            this.setState({ showMessageBar: false, disableBtn: false });
            return;
        }
        await this.loadEngagements(this.state.TeamURL);
        this.setState({ showMessageBar: false });

    }

    public CheckSplitRollover = () => {

        let currentengNumber = "";
        let newEngagementNumber = "";
        if (Isnextyear == true) {
            let updatedYear = this.state.Year.toString().slice(-2);
            let engNumber = this.state.EngagementNumberSelected.substring(0, this.state.EngagementNumberSelected.length - 2);
            newEngagementNumber = engNumber + updatedYear;
        }

        if (this.state.EngagementNumberSelected.toString().slice(-2) == "00" && this.state.EngagementNumberSelected.lastIndexOf("-00") > -1) {
            currentengNumber = newEngagementNumber;
        }
        else {
            currentengNumber = this.state.EngagementNumberSelected;
        }

        let SplitRollover = new AssuranceSplitRollover();
        let eng = currentengNumber.slice(-2);
        let e1 = parseInt(eng) - 1;
        let str1 = currentengNumber.slice(0, -2) + e1.toString();
        SplitRollover.GetAssuranceSplitRollover(str1, currentengNumber, this.state.TeamSelected, this.state.PortalTypeSelected, Isnextyear).then((results) => {
            this.setState({
                AssuranceSplitRollover: results
            });
        });
    }

    public Rollover = async () => {
        let PortalType = this.state.PortalTypeSelected;
        let Team = this.state.TeamSelected;
        let hubWeb = Web(GlobalValues.HubSiteURL);
        if (Isnextyear == true) {
            await hubWeb.lists.getByTitle(GlobalValues.EngagementPortalList).items.filter("EngagementNumberEndZero eq '" + this.state.EngagementNumberSelected + "'").getAll().then((data) => {
                data = data.filter(e => e.PortalExist == true && e.ClientNumber == CRN && e.PortalType == PortalType && e.Team == Team);
                let eng = this.state.UpdatedEngagementNumberSelected.slice(-2);
                let e1 = parseInt(eng) - 1;
                let str1 = this.state.UpdatedEngagementNumberSelected.slice(0, -2) + e1.toString();
                this.setState({
                    Rollover: true,
                    ServiceTypeSelected: data[0].TemplateType,
                    IndustryTypeSelected: data[0].IndustryType,
                    SupplementalSelected: data[0].Supplemental,
                    RolloverURL: str1
                });
                this.state.CRUserList.forEach((e) => {

                    if (data[0].EngagementMembers.indexOf(e.email) > -1) {
                        e.checked = true;
                    }
                });

                this.state.CLUserList.forEach((e) => {

                    if (data[0].ClientMembers.indexOf(e.email) > -1) {
                        e.checked = true;
                    }
                });

                if (this.state.AssuranceSplitRollover.length == 0) {
                    this._getUserList();
                }
            });
        }
        else {
            let eng = this.state.EngagementNumberSelected.slice(-2);
            let e1 = parseInt(eng) - 1;
            let str1 = this.state.EngagementNumberSelected.slice(0, -2) + e1.toString();

            await hubWeb.lists.getByTitle(GlobalValues.EngagementPortalList).items.filter("Title eq '" + str1 + "'").getAll().then((data) => {
                data = data.filter(e => e.PortalExist == true && e.ClientNumber == CRN && e.PortalType == PortalType && e.Team == Team);
                if (data.length != 0) {
                    let WorkYear = parseInt(data[0].WorkYear);
                    let Year = parseInt(this.state.Year);
                    if (Year == WorkYear + 1) {
                        this.setState({
                            Rollover: true,
                            ServiceTypeSelected: data[0].TemplateType,
                            IndustryTypeSelected: data[0].IndustryType,
                            SupplementalSelected: data[0].Supplemental,
                            RolloverURL: str1
                        });

                        this.state.CRUserList.forEach((e) => {
                            if (data[0].EngagementMembers.indexOf(e.email) > -1) {
                                e.checked = true;
                            }
                        });

                        this.state.CLUserList.forEach((e) => {
                            if (data[0].ClientMembers.indexOf(e.email) > -1) {
                                e.checked = true;
                            }
                        });

                        if (this.state.AssuranceSplitRollover.length == 0) {
                            this._getUserList();
                        }
                    }
                    else {
                        this.setState({ isRollover: true });
                    }
                } else {
                    this.setState({ isRollover: true });
                }
            });
        }

    }

    public _onChangePortalChoice = (event: React.FormEvent<HTMLDivElement>, option: IChoiceGroupOption) => {
        this._getUserListCreatedon();
        this.setState({ PortalChoiceSelected: option.text });
        let ErrorMessage = "";
        this.setState({ Message: ErrorMessage, showMessageBar: false, MessageBarType: OfficeUI.MessageBarType.error, disableBtn: false });
        if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Advisory") {
            this._getUserListAdvisory();
        }
    }
    private closeMessageBar = () => {
        this.setState({ showMessageBar: false });
    }

    private _getPeoplePickerItems(items: any[]) {
        const currSite = Web(GlobalValues.HubSiteURL);
        let getSelectedUsers = [];
        let getusersEmails = [];
        for (let item in items) {
            getSelectedUsers.push(items[item].text);
            getusersEmails.push(items[item].secondaryText);
        }
        items.forEach((e) => {
          currSite.siteUsers.getByLoginName(e.loginName).get().then((user) => {
            this.setState({ addusers: getSelectedUsers, addusersID: user.Id, emailaddress: getusersEmails });
            });
        });
    }

    private _validateSiteOwner(items: any[]) {
        // show error message if this is a guest user
        let userEmail = items[0].secondaryText.toLowerCase();
        if ((userEmail.indexOf('cohnreznick.com') == -1) && (userEmail.indexOf('cohnreznickdev') == -1)) {
            // this is a guest user, do not validate
        }
        else {
            this._getPeoplePickerItems(items);
        }
    }

    // validate the user is a CR user:
    private _validateEngagementMembers(items: any[]) {
        this.setState({validate: false});
        let validateEmails = true;
        // show error message if this is a guest user
        items.forEach((e) => {
            let userEmail = e.secondaryText.toLowerCase();
            if ((userEmail.indexOf('cohnreznick.com') == -1) && (userEmail.indexOf('cohnreznickdev') == -1)) {
                validateEmails = false;
                console.log('show error');
            }
        });
        if (validateEmails == true) {
            this._getUserItems(items);
        }
        else {
            this.setState({validate: true});
        }
    }

     // NEW People Picker for adding users.  Per Converge team, only CR users can be added at this time
     private async _getUserItems(items: any[]) {
        let selectedUsers = [];
        let accessUserList = [];
        items.forEach((e) => {
            accessUserList.push(e.secondaryText);
            selectedUsers.push(e.text);
        });
        this.setState({AccessUserList: accessUserList, addusers1: selectedUsers});
    }

    private getCLUserList() {
        let CLUserSelected = '';
        let CRUserSelected = '';
        let FinalAccessUserList = '';
        this.state.CLUserList.forEach((e) => {
            if (e.checked) {
                CLUserSelected += e.email + ";";
            }
        });

        this.state.CRUserList.forEach((e) => {
            if (e.checked) {
                CRUserSelected += e.email + ";";
            }
        });

        this.state.AccessUserList.forEach((e) => {
            FinalAccessUserList += e + ";";
        });

        this.setState({
            CLUserSelected: CLUserSelected,
            CRUserSelected: CRUserSelected,
            FinalAccessUserList: FinalAccessUserList
        });
    }

    private _onSelectDate = (date: Date | null | undefined): void => {
        this.setState({
            portalExpiration: date
        });
    }

    private _onSelectDate2 = (date: Date | null | undefined): void => {
        this.setState({
           // DateExtend: date
           portalExpiration: portalExpDate,
        });
    }

    private _onSelectDate3 = (date: Date | null | undefined): void => {
        this.setState({
            K1Date: date
        });
    }

    private _onSelectDateFileExp = (date: Date | null | undefined): void => {
        portalExpDate = addMonths(date, 6); // set the portal expiration for 6 months after the file expiration
        this.setState({
            // DateExtend: date
            portalExpiration: portalExpDate,
            fileExpiration: date // this is the date the user picked
        });
    }

    private _onFormatDate = (date: Date): string => {
        return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
    }


    private OnFileSelect = () => {
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        this.setState({ K1FileName: myfile.name });

    }

    private _getUserListCreatedon() {
        let obj = new ClientInfoClass();
        //let userlist = "";
        let userlist = [];
        obj.GetUsersByGroup("CL-" + CRN
        ).then((results) => {
            results.forEach((e) => {
                //this.state.AccessUserList.push({ name: e.Email });
                userlist.push({ name: e.Email });
                //userlist += e.Email + ";";
            });
            //this.setState({ FinalAccessUserList: userlist });
            this.setState({ AccessUserList: userlist });
        });
    }

    private async _getUserListAdvisory() {
        let obj = new ClientInfoClass();
        await obj.GetUsersByGroup("CRAD-ADV-" + CRN).then((res) => {
            res.forEach((e) => {
                this.state.CLUserList.push({ email: e.Email, checked: false });
            });
        });
    }

    private async _getUserList() {
        try {
            let obj = new ClientInfoClass();
            if (this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow') {
                await obj.GetUsersByGroup("CRET-TAX-WF-" + this.state.RolloverURL).then((res) => {
                    res.forEach((e) => {
                        this.state.CRUserList.push({ email: e.Email, checked: false });
                    });
                });

                await obj.GetUsersByGroup("CL-TAX-WF-" + this.state.RolloverURL).then((res) => {
                    res.forEach((e) => {
                        this.state.CLUserList.push({ email: e.Email, checked: false });
                    });
                });

            }
            else if (this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow') {

                await obj.GetUsersByGroup("CRET-AUD-WF-" + this.state.RolloverURL).then((res) => {
                    res.forEach((e) => {
                        this.state.CRUserList.push({ email: e.Email, checked: false });
                    });
                });

                await obj.GetUsersByGroup("CL-AUD-WF-" + this.state.RolloverURL).then((res) => {
                    res.forEach((e) => {
                        this.state.CLUserList.push({ email: e.Email, checked: false });
                    });
                });
            }
        }
        catch {
            return console.warn;
        }

    }


    public onChangeEmailCRList = (value, email) => {
        console.log('firing onChangeEmailCRList');
        // console.log('logging user email: ', email);

        let CRList = this.state.CRUserList;

        CRList.forEach((e) => {
            if (e.email == email && value) {
                e.checked = true;

            }

            if (e.email == email && !value) {
                e.checked = false;
            }
        });
        this.setState({ CRUserList: CRList });
        // this.setState({ PreExistingAlertUsers: checkedUsers });
        this.formulatePreExistingAlertUsers();
    }

    public onChangeEmailCLList = (value, email) => {
        let CLList = this.state.CLUserList;
        CLList.forEach((e) => {
            if (e.email == email && value) {
                e.checked = true;
                // checkedUsers.push(e);
            }

            if (e.email == email && !value) {
                e.checked = false;
            }
        });
        this.setState({ CLUserList: CLList });
        this.formulatePreExistingAlertUsers();

    }


    public formulatePreExistingAlertUsers = async () => {
      console.log('formulatePreExistingAlertUsers firing::');
      const prevUsersToRollAlertsState = this.state.UsersToRollAlerts;
      const checkedCRUsers = [];
      const checkedCLUsers = [];
      let allCheckedUsers = [];
      let usersWithAlerts = [];
      const portalWeb = Web(GlobalValues.SiteURL);

      this.state.CRUserList.forEach(e => {
        if (e.checked) {
          checkedCRUsers.push(e);
        }
      });

      this.state.CLUserList.forEach(e => {
        if (e.checked) {
          checkedCLUsers.push(e);
        }
      });

      allCheckedUsers = [...checkedCRUsers, ...checkedCLUsers];

      const filteredObjs = prevUsersToRollAlertsState.filter(obj1 => {
        return allCheckedUsers.some(obj2 => {
          return obj1.email === obj2.email;
        });
      });

      // use for of loop so we can await on the api calls
      for (const user of allCheckedUsers) {
        const userObj = await portalWeb.siteUsers.getByEmail(user.email).get();
        console.log(userObj.Id);
        // send api call to retrieve alerts for rolloverURL based on userId
        const userAlertData = await fetch(`${GlobalValues.SiteURL + "/" + this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + this.state.RolloverURL}/_api/web/alerts?$filter=UserId eq ${userObj.Id}`,
            {
              headers: {
                Accept: "application/json;odata=verbose",
              },
            })
            .then(data => {
              return data.json();
            });

        console.log('logging userHasAlert: ', userAlertData);

        if (userAlertData.d.results.length) {
          // usersWithAlerts.push(user);
          user.hasAlert = true
        } else {
          user.hasAlert = false;
        }

      }



      console.log('loggign usersWithAlerts before setting state: ', allCheckedUsers);
      // console.log('loggign usersWithAlerts before setting state: ', usersWithAlerts);

      // console.log('setting PreExistingAlertUsers state::');
      this.setState({ PreExistingAlertUsers: allCheckedUsers });
      this.setState({ UsersToRollAlerts: filteredObjs });
    }





    public onChangeUsersToRollAlerts = (val, email) => {
      console.log('onChangeUsersToRollAlerts firing:::');
      let output = [];

      const prevInfoState = this.state.UsersToRollAlerts;
      const preExistingAlertUsersHolder = this.state.PreExistingAlertUsers;

      const filteredObjs = prevInfoState.filter(obj1 => {
        return preExistingAlertUsersHolder.some(obj2 => {
          return obj1 !== obj2;
        });
      });


      let tempUserArr1 = [];
      let tempUserArr2 = [];

      if (val) {
        tempUserArr1.push({ email: email, checkedState: val });
        tempUserArr2 = [...filteredObjs, ...tempUserArr1];
        this.setState({ UsersToRollAlerts: tempUserArr2 });
      } else {
        output = filteredObjs.filter(value => {
          return value.email !== email;
        });
        tempUserArr2 = output;
        this.setState({ UsersToRollAlerts: tempUserArr2 });
      }

      preExistingAlertUsersHolder.forEach((e) => {
        if (e.email == email && val) {
            e.checkedState = true;
        }

        if (e.email == email && !val) {
            e.checkedState = false;
        }
      });

      this.setState({ PreExistingAlertUsers: preExistingAlertUsersHolder });

    }


    public GetYearOption = () => {
        let date = new Date();
        let year = date.getFullYear();
        let option = [];

        for (let i = year - 3; i <= year + 3; i++) {
            option.push({
                key: i.toString(),
                text: i.toString()
            });
        }
        return option;
    }

    public onChangeYear = (event, item) => {
        if (item) {
            this.setState({ Year: item.key });

            if (updatedworkyear == true) {
                Isnextyear = false;
                let hubWeb = Web(GlobalValues.HubSiteURL);
                hubWeb.lists.getByTitle(GlobalValues.EngagementPortalList).items.filter("EngagementNumberEndZero eq '" + this.state.EngagementNumberSelected + "'").getAll().then((data) => {
                    let data1 = data.filter(e => e.WorkYear == item.key && e.ClientNumber == CRN && e.PortalType == this.state.PortalTypeSelected && e.Team == this.state.TeamSelected);
                    let data2 = data.filter(e => e.WorkYear == parseInt(item.key) - 1 && e.ClientNumber == CRN && e.PortalType == this.state.PortalTypeSelected && e.Team == this.state.TeamSelected);
                    let data3 = data.filter(e => e.WorkYear == parseInt(item.key) - 1 && e.ClientNumber == CRN && e.PortalType == this.state.PortalTypeSelected && e.Team == this.state.TeamSelected && e.SplitSuffix != "");
                    if (data1.length > 0) {
                        let ErrorMessage = "You can not create same engagement number " + this.state.EngagementNumberSelected + " for " + this.state.PortalTypeSelected + " (Year:" + item.key + ")";
                        this.setState({ Message: ErrorMessage, showMessageBar: true, MessageBarType: OfficeUI.MessageBarType.error, Checkeng: false, disableBtn: true, });
                        return false;
                    }
                    else {
                        if (data2.length == 1) {
                            let rolloveryear = parseInt(data2[0].WorkYear) + 1;
                            if (rolloveryear == item.key) {
                                Isnextyear = true;
                            }
                            else {
                                Isnextyear = false;
                            }
                        }
                        else if (data3.length >= 1) {
                            let rolloveryear = parseInt(data3[0].WorkYear) + 1;
                            if (rolloveryear == item.key) {
                                Isnextyear = true;
                            }
                            else {
                                Isnextyear = false;
                            }
                        }

                        let ErrorMessage = "";
                        this.setState({ Message: ErrorMessage, showMessageBar: false, MessageBarType: OfficeUI.MessageBarType.error, Checkeng: true, disableBtn: false });
                        return true;
                    }
                });
            }

            if (this.state.PortalTypeSelected == "K1") {
                ExDate = (6) + '-' + (1) + '-' + (parseInt(item.key) + 2);
                let dt = new Date(ExDate);
                const ExDate1: Date = dt;
                this.setState({
                    K1Date: ExDate1,
                });
            }
        }

    }

    public CloseButton = () => {
        this.ResetState();
        this.setState({
            isOpen: false,
            currentScreen: ""
        });
        window.location.reload();
    }


    public onItemSelected = (item: ITag): ITag | null => {
        if (item && item.name) {
            EngagementNameTags = [{ key: item.key.toString(), name: item.name }];
        }
        return item;
    }

    public render(): React.ReactElement<ICreateEngagement> {

        const inputProps: IInputProps = {
            onBlur: (ev: React.FocusEvent<HTMLInputElement>) => { console.log('onBlur called'); },
            onFocus: (ev: React.FocusEvent<HTMLInputElement>) => { console.log('onFocus called'); },
            'aria-label': 'Tag picker',
        };

        const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
            suggestionsHeaderText: 'Suggested tags',
            noResultsFoundText: 'No Engagement found',
        };

        const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
            if (!tagList || !tagList.length || tagList.length === 0) {
                return false;
            }
            return tagList.some(compareTag => compareTag.key === tag.key);
        };

        const filterSuggestedTags1 = (filterText: string, tagList: ITag[]): ITag[] => {
            if (EngagementNumberTags != undefined) {
                return filterText
                    ? EngagementNumberTags.filter(
                        tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 && !listContainsTagList(tag, tagList),
                    )
                    : [];
            }
        };

        const getTextFromItem1 = (item: ITag) => item.name;

        let title = "Create Engagement Subportal " + this.state.titleText;
        let K1MaxDate = new Date(ExDate);

        let HubSiteURL = GlobalValues.HubSiteURL;
        let K1ExcelTemplate = "/SiteAssets/K1ExcelTemplate.csv";
        let TagpickerDisable = true;
        if (this.state.TeamSelected != "" && this.state.PortalTypeSelected != "") {
            TagpickerDisable = false;
        }

        return (
            <div>
                <Link className={styles.primarybutton} onClick={(e) => this.openDialog(e)}>Create Engagement Subportal</Link>
                <Dialog
                    isOpen={this.state.isOpen}
                    onDismiss={close.bind(this)}
                    isBlocking={true}
                    containerClassName={styles.createEngagement}
                    dialogContentProps={{
                        title: title
                    }}
                >
                    <div>
                        <div className={contentStyles.body}>
                            <ProgressBar isVisible={this.state.showSpinner} Message="Please Wait"></ProgressBar>
                            {this.state.currentScreen == "screen1" ?
                                <div className={styles.screenOne}>
                                    <Stack horizontal gap={20}>
                                        <ChoiceGroup
                                            className={styles.innerChoice}
                                            options={
                                                this.state.TeamSelected == "Tax" ? PortalTypeOptionsForTax : PortalTypeOptions
                                            }
                                            label="Portal Type"
                                            onChange={this._onChangePortalType}
                                            required={true}
                                            selectedKey={this.state.PortalTypeSelected}
                                        />
                                    </Stack>
                                    <div className={styles.innerChoiceDesc}>
                                        <div className={styles.choiceDes}>
                                            <text>This is a designated worksite area for engagement workflow functionality.</text>
                                        </div>
                                        <div className={styles.choiceDes}>
                                            <text>File sharing collaboration platform designated for document management.</text>
                                        </div>
                                        <div className={styles.choiceDes}>
                                            <text>Custom section for the delivery of K1 documents to clients of CR clients.</text>
                                        </div>
                                    </div>
                                    {(this.state.validate && this.state.PortalTypeSelected == "") ?
                                        <div className={styles.reqval}>Portal Type is mandatory.</div> : ''}

                                    {this.state.PortalTypeSelected != 'K1' ?
                                        <div className={styles.teamRadio}>
                                            <Stack horizontal gap={20}>
                                                <ChoiceGroup
                                                    className={styles.innerChoice}
                                                    options={this.state.PortalTypeSelected == "File Exchange" ? TeamoptionsFileExchange : Teamoptions}
                                                    label="Team"
                                                    required={true}
                                                    onChange={this._onChangeTeam}
                                                    selectedKey={this.state.TeamSelected}
                                                />
                                            </Stack>
                                            {(this.state.validate && this.state.TeamSelected == "") ?
                                                <div className={styles.reqval}>Team is mandatory.</div> : ''}
                                        </div>

                                        : <div className={styles.teamRadio}>
                                            <Stack horizontal gap={20}>
                                                <ChoiceGroup
                                                    className={styles.innerChoice}
                                                    options={TeamoptionsK1}
                                                    label="Team"
                                                    selectedKey="Tax"
                                                />
                                            </Stack>
                                            {(this.state.validate && this.state.TeamSelected == "") ?
                                                <div className={styles.reqval}>Team is mandatory.</div> : ''}
                                        </div>}
                                    <div className={styles.engnumbername}>
                                        <div className={`${styles.engagementnames} ${styles.column1}`}>
                                            <Label>Engagement Number<span className={styles.reqval}> *</span></Label>
                                            <TooltipHost content="Enter Engagement Number">
                                                <TagPicker
                                                    defaultSelectedItems={EngagementNameTags}
                                                    removeButtonAriaLabel="Remove"
                                                    onResolveSuggestions={filterSuggestedTags1}
                                                    getTextFromItem={getTextFromItem1}
                                                    pickerSuggestionsProps={pickerSuggestionsProps}
                                                    itemLimit={1}
                                                    inputProps={inputProps}
                                                    onChange={this._onChangeEngagementNumber}
                                                    disabled={TagpickerDisable}
                                                    onItemSelected={this.onItemSelected}
                                                />
                                            </TooltipHost>
                                            {(this.state.validate && this.state.EngagementNumberSelected == "") ?
                                                <div className={styles.reqval}>Invalid Engagement Number. Please enter a correct Engagement Number and try again.
                                                </div> : ''}
                                        </div>
                                        <div className={`${styles.engagementnames} ${styles.column2}`}>
                                            <Label>Engagement Name</Label>
                                            <TooltipHost
                                                content="Enter Engagement Name">
                                                <TextField disabled className={styles.engagementPrint} defaultValue={this.state.EngagementNameSelected} />
                                            </TooltipHost>
                                        </div>
                                    </div>
                                    <div className={styles.engnumbername}>
                                        <div className={`${styles.engagementnames} ${styles.column1}`}>
                                            <Label>Year <span className={styles.reqval}> *</span></Label>
                                            {updatedworkyear == true ?
                                                <Dropdown
                                                    placeholder="Select an option"
                                                    onChange={this.onChangeYear}
                                                    options={this.GetYearOption()}
                                                    selectedKey={this.state.Year}
                                                />
                                                : <TextField disabled className={styles.engagementPrint} defaultValue={this.state.Year} />
                                            }
                                            {(this.state.validate && (this.state.Year == "" || this.state.Year == null)) ?
                                                <div className={styles.reqval}>Year is mandatory.</div> : ''
                                            }
                                        </div>
                                        <div className={`${styles.engagementnames} ${styles.column2}`}>
                                            <PeoplePicker
                                                context={this.props.spContext}
                                                titleText="Site Owner"
                                                showtooltip={false}
                                                required={true}
                                                onChange={(items) => this._validateSiteOwner(items)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User]}
                                                ensureUser={true}
                                                personSelectionLimit={1}
                                                placeholder="Enter name or email"
                                                defaultSelectedUsers={this.state.addusers}
                                            />
                                            {(this.state.validate && this.state.addusers.length == 0) ?
                                                <div className={styles.reqval}>Site Owner is mandatory and must be a CohnReznick employee.</div> : ''
                                            }
                                        </div>
                                    </div>
                                    {this.state.showMessageBar && <OfficeUI.MessageBar
                                        messageBarType={this.state.MessageBarType}
                                        isMultiline={false}
                                        onDismiss={() => this.closeMessageBar()}
                                        dismissButtonAriaLabel="Close">

                                        {this.state.Message}
                                    </OfficeUI.MessageBar>}
                                </div> : ""}

                            {this.state.currentScreen == 'screen2' ?
                                <div className={styles.screenTwo}>

                                    {this.state.showMessageBar && <OfficeUI.MessageBar
                                        messageBarType={this.state.MessageBarType}
                                        isMultiline={false}
                                        onDismiss={() => this.closeMessageBar()}
                                        dismissButtonAriaLabel="Close">
                                        {this.state.Message}
                                    </OfficeUI.MessageBar>}
                                    <div className={styles.freshRollover}>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Portal Type</Label>
                                                <Text className={styles.engagementPrint}>{this.state.PortalTypeSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Team</Label>
                                                <Text className={styles.engagementPrint}>{this.state.TeamSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Number</Label>
                                                <Text className={styles.engagementPrint}>
                                                    {updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Name</Label>
                                                <Text className={styles.engagementPrint}>{this.state.EngagementNameSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Year</Label>
                                                <Text className={styles.engagementPrint}>{this.state.Year}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Site Owner</Label>
                                                <Text className={styles.engagementPrint}>{this.state.addusers}</Text>
                                            </div>
                                        </div>
                                    </div>
                                    <div className={styles.divider}></div>
                                    <div className={styles.engagementnames}>
                                        <Label>{this.state.PortalChoiceSelected} Portal<span className={styles.reqval}> *</span></Label>
                                        {this.state.PortalChoiceSelected == 'Rollover' ?
                                            <Text className={styles.engagementPrint}>All requests from previous year's portal will be rolled over to this portal on creation.</Text> :
                                            <Text className={styles.engagementPrint}>A portal will be set up with new template and requests.</Text>
                                        }
                                    </div>
                                    <Stack horizontal gap={20} className="portalChoice">
                                        <ChoiceGroup
                                            className={styles.innerChoice}
                                            options={this.state.isRollover ? PortalChoiceOptions1 : PortalChoiceOptions}
                                            onChange={this._onChangePortalChoice}
                                            selectedKey={this.state.PortalChoiceSelected}
                                        />

                                    </Stack>
                                    <div className={styles.innerChoiceDesc}>
                                        <div className={styles.choiceDes}>
                                            <text>This option will rollover all requests from the previous year's portal.</text>
                                        </div>
                                        <div className={styles.choiceDes}>
                                            <text>This option will allow you to choose a new template and start with an empty portal.</text>
                                        </div>
                                    </div>
                                    {(this.state.validate && this.state.PortalChoiceSelected == "") ?
                                        <div className={styles.reqval}>Portal selection is required </div> : ''}
                                </div> : ""}

                            {this.state.currentScreen == 'screen3' ?
                                <div className={styles.screenThree}>
                                    <div className={styles.freshRollover}>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Portal Type</Label>
                                                <Text className={styles.engagementPrint}>{this.state.PortalTypeSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Team</Label>
                                                <Text className={styles.engagementPrint}>{this.state.TeamSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Number</Label>
                                                <Text className={styles.engagementPrint}>
                                                    {updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Name</Label>
                                                <Text className={styles.engagementPrint}>{this.state.EngagementNameSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Year</Label>
                                                <Text className={styles.engagementPrint}>{this.state.Year}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Site Owner</Label>
                                                <Text className={styles.engagementPrint}>{this.state.addusers}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            {this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Rollover' ?
                                                <div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Template Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.ServiceTypeSelected}</Text>
                                                    </div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Industry Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.IndustryTypeSelected}</Text>
                                                    </div>
                                                    <div className={styles.visualhidden}>
                                                        <Label>.</Label>
                                                        <Text className={styles.engagementPrint}>.</Text>
                                                    </div>
                                                </div> : ""}
                                        </div>
                                    </div>
                                    <div className={styles.divider}></div>
                                    <div className={styles.serviceindustryType}>
                                        {this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                            <div className={styles.taxType}>
                                                <Dropdown
                                                    placeholder="Template Type"
                                                    label="Template Type"
                                                    options={this.state.ServiceType}
                                                    onChange={this._onChangeServiceType}
                                                    required={true}
                                                    selectedKey={this.state.ServiceTypeSelectedKey}
                                                />
                                                {(this.state.validate && this.state.ServiceTypeSelected == "") ?
                                                    <div className={styles.reqval}>Template Type is required </div> : ''}
                                                <Dropdown
                                                    placeholder="Industry Type"
                                                    label="Industry Type"
                                                    options={this.state.IndustryType}
                                                    onChange={this._onChangeIndustryType}
                                                    required={true}
                                                    selectedKey={this.state.IndustryTypeSelectedKey}
                                                />
                                                {(this.state.validate && this.state.IndustryTypeSelected == "") ?
                                                    <div className={styles.reqval}>Industry Type is required </div> : ''}
                                            </div> : ""}
                                        {this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                            <div className={styles.assuranceType}>
                                                <Label>Selecting an Industry Type will generate a fully populated set of industry-specific request items. If you'd prefer a blank template, please select "N/A".</Label>
                                                <div className={styles.Supplemental}>
                                                    <Dropdown
                                                        placeholder="Industry Type"
                                                        label="Industry Type"
                                                        options={this.state.IndustryType}
                                                        onChange={this._onChangeIndustryType}
                                                        required={true}
                                                        selectedKey={this.state.IndustryTypeSelectedKey}
                                                    />
                                                    {(this.state.validate && this.state.IndustryTypeSelected == "") ?
                                                        <div className={styles.reqval}>Industry Type is required </div> : ''}
                                                    <Dropdown
                                                        placeholder="Supplemental"
                                                        label="Supplemental"
                                                        options={this.state.Supplemental}
                                                        onChange={this._onChangeSupplemental}
                                                        required={true}
                                                        selectedKey={this.state.SupplementalSelectedKey}
                                                    />
                                                    {(this.state.validate && this.state.SupplementalSelected == "") ?
                                                        <div className={styles.reqval}>Supplemental is required </div> : ''}
                                                </div>
                                            </div> : ""}
                                        {this.state.TeamSelected == 'Advisory' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                            <div className={styles.advisoryType}>
                                                <Dropdown
                                                    placeholder="Template Type"
                                                    label="Template Type"
                                                    options={this.state.AdvisoryTemplate}
                                                    onChange={this._onChangeAdvisoryTemplate}
                                                    required={true}
                                                    selectedKey={this.state.AdvisoryTemplateSelectedKey}
                                                />
                                                {(this.state.validate && this.state.AdvisoryTemplateSelected == "") ?
                                                    <div className={styles.reqval}>Template Type is required </div> : ''}
                                            </div> : ""}
                                    </div>
                                    {this.state.TeamSelected == 'Assurance' ? <div className={styles.workpaper}>
                                        <Label className={styles.wlabeltxt}>Enter the workpaper path of the binder in CCH Engagement where the documents should be automatically exported. If you leave this field blank, the documents will not be automatically uploaded to CCH Engagement.</Label>
                                        <TextField label="Workpaper Path" value={this.state.WorkpaperPath} onChange={(ev: React.FormEvent<HTMLElement>, newValue?: string) => (this.setState({ WorkpaperPath: newValue }))} />
                                    </div> : ""}
                                    {/* Do NOT DELETE THIS CODE */}
                                    {/* Do NOT DELETE THIS CODE */}
                                    {/* Do NOT DELETE THIS CODE */}
                                    {/* {this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                        <AssuranceEngSplit OnSplitChange={this.SetAssuranceSplitData} AsuranceSplitData={this.state.AsuranceSplitData} spContext={this.props.spContext}></AssuranceEngSplit>
                                        : ""} */}
                                    <div>
                                        <div className={styles.formcontrol}>
                                            {this.state.PortalChoiceSelected == 'Create New' ?
                                                <div>
                                                    {this.state.TeamSelected == 'Advisory' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                                        <div>
                                                            <Label>Please select which users should have access to this portal:</Label>
                                                            <div className={styles.usergroups}>
                                                                {this.state.CLUserList.filter(element => element.email !== "").map(element =>
                                                                    <Checkbox label={element.email} checked={element.checked} onChange={(ev, value) => {
                                                                        this.onChangeEmailCLList(value, element.email);
                                                                    }} />
                                                                )
                                                                }
                                                            </div>
                                                        </div>
                                                        : ""
                                                    }
                                                </div>
                                                : ""}

                                                {/* */}
                                                {/*  */}
                                            {this.state.PortalChoiceSelected == 'Rollover' ?
                                                <div>
                                                    <div>
                                                        <Label>The following users will automatically have access:</Label>
                                                        {this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow' ?
                                                        <>
                                                            <div className={styles.userLists}>
                                                                {
                                                                    <div className={styles.usergroups}>
                                                                        CRET-TAX-WF-{this.state.EngagementNumberSelected}
                                                                        {this.state.CRUserList.filter(element => element.email !== "").map(element =>
                                                                            <Checkbox label={element.email} checked={element.checked} onChange={(ev, value) => {
                                                                                this.onChangeEmailCRList(value, element.email);
                                                                            }} />
                                                                        )
                                                                        }
                                                                    </div>
                                                                }
                                                                {
                                                                    <div className={styles.usergroups}>
                                                                        CL-TAX-WF-{this.state.EngagementNumberSelected}
                                                                        {this.state.CLUserList.filter(element => element.email !== "").map(element =>
                                                                            <Checkbox label={element.email} checked={element.checked} onChange={(ev, value) => {
                                                                                this.onChangeEmailCLList(value, element.email);
                                                                            }} />
                                                                        )
                                                                        }
                                                                    </div>
                                                                }
                                                            </div>
                                                            { this.state.PreExistingAlertUsers.filter(e => e.hasAlert === true).length > 0 &&
                                                              <div className={styles.userLists}>
                                                                <Label>Select the users to rollover alerts for:</Label>
                                                              <div className={styles.usergroups}>
                                                                 {/* {console.log('logging PreExistingAlertUsers: ', this.state.PreExistingAlertUsers)} */}
                                                                  {this.state.PreExistingAlertUsers.filter(e => e.hasAlert === true).map(element =>
                                                                    <Checkbox label={element.email} checked={element.checkedState} onChange={(ev, value) => {
                                                                      this.onChangeUsersToRollAlerts(value, element.email);
                                                                  }} />
                                                                  )}
                                                                  {/* {console.log('logging UsersToRollAlerts: ', this.state.UsersToRollAlerts)} */}
                                                                  </div>
                                                              </div>
                                                            }
                                                            </> :
                                                            this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.AssuranceSplitRollover.length == 0 ?
                                                            <>
                                                                <div className={styles.userLists}>
                                                                    <div className={styles.usergroups}>
                                                                        CRET-AUD-WF-{this.state.EngagementNumberSelected}
                                                                        {this.state.CRUserList.filter(element => element.email !== "").map(element =>
                                                                            <Checkbox label={element.email} checked={element.checked} onChange={(ev, value) => {
                                                                                this.onChangeEmailCRList(value, element.email);
                                                                            }} />
                                                                        )
                                                                        }
                                                                    </div>
                                                                    <div className={styles.usergroups}>
                                                                        CL-AUD-WF-{this.state.EngagementNumberSelected}
                                                                        {this.state.CLUserList.filter(element => element.email !== "").map(element =>
                                                                            <Checkbox label={element.email} checked={element.checked} onChange={(ev, value) => {
                                                                                this.onChangeEmailCLList(value, element.email);
                                                                            }} />
                                                                        )
                                                                        }
                                                                    </div>
                                                                </div>
                                                                {this.state.PreExistingAlertUsers.filter(e => e.hasAlert === true).length > 0 &&
                                                              <div className={styles.userLists}>
                                                                <Label>Select the users to rollover alerts for:</Label>
                                                              <div className={styles.usergroups}>
                                                                {/* {console.log('logging PreExistingAlertUsers: ', this.state.PreExistingAlertUsers)} */}
                                                                  {this.state.PreExistingAlertUsers.filter(e => e.hasAlert === true).map(element =>
                                                                      <Checkbox label={element.email} checked={element.checkedState} onChange={(ev, value) => {
                                                                        this.onChangeUsersToRollAlerts(value, element.email)
                                                                      }} />
                                                                  )}
                                                                  {/* {console.log('logging UsersToRollAlerts: ', this.state.UsersToRollAlerts)} */}
                                                                  </div>
                                                              </div>
                                                        }
                                                            </> : ""}
                                                    </div>
                                                    {/* Do NOT DELETE THIS CODE */}
                                                    {/* Do NOT DELETE THIS CODE */}
                                                    {/* Do NOT DELETE THIS CODE */}
                                                    {/*
                                                    {this.state.TeamSelected == 'Assurance' && this.state.AssuranceSplitRollover.length > 0 && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Rollover' ?
                                                        <AssuranceEngSplitRolloverUsers spContext={this.props.spContext} Data={this.state.AssuranceSplitRollover} Control={this.SetAssuranceSplitDataRollOver}></AssuranceEngSplitRolloverUsers>
                                                        : ""}

                                                    <div className={styles.divider}></div>

                                                    <div>
                                                        {this.state.TeamSelected == 'Assurance' && this.state.AssuranceSplitRollover.length > 0 && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Rollover' ?

                                                            <AssuranceEngSplitRollover spContext={this.props.spContext} Data={this.state.AssuranceSplitRollover} Control={this.SetAssuranceSplitDataRollOver}></AssuranceEngSplitRollover>
                                                            : ""}
                                                    </div> */}

                                                </div> : ""}
                                        </div>
                                        <div className={styles.pnppicker}>
                                            <PeoplePicker
                                                context={this.props.spContext}
                                                titleText={this.state.peoplePickerTitle}
                                                showtooltip={false}
                                                required={false}
                                                disabled={false}
                                                onChange={(items) => this._validateEngagementMembers(items)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User]}
                                                ensureUser={true}
                                                personSelectionLimit={100}
                                                placeholder="Enter name or email"
                                                defaultSelectedUsers={this.state.addusers1}
                                            />
                                            <span className={styles.optional}>optional</span>
                                             {this.state.validate ?
                                                <div className={styles.reqval}>Users added here must be CohnReznick employees.</div> : ''
                                            }
                                        </div>
                                    </div>
                                    {this.state.showMessageBar && <OfficeUI.MessageBar
                                        messageBarType={this.state.MessageBarType}
                                        isMultiline={false}
                                        onDismiss={() => this.closeMessageBar()}
                                        dismissButtonAriaLabel="Close">
                                        {this.state.Message}
                                    </OfficeUI.MessageBar>}
                                </div> : ""}
                            {this.state.currentScreen == 'screen4' ?
                                <div className={styles.screenFour}>
                                    <div className={styles.freshRollover}>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Portal Type</Label>
                                                <Text className={styles.engagementPrint}>{this.state.PortalTypeSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Team</Label>
                                                <Text className={styles.engagementPrint}>{this.state.TeamSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Number</Label>
                                                <Text className={styles.engagementPrint}>
                                                    {updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Name</Label>
                                                <Text className={styles.engagementPrint}>{this.state.EngagementNameSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Year</Label>
                                                <Text className={styles.engagementPrint}>{this.state.Year}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Site Owner</Label>
                                                <Text className={styles.engagementPrint}>{this.state.addusers}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            {this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow' ?
                                                <div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Template Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.ServiceTypeSelected}</Text>
                                                    </div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Industry Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.IndustryTypeSelected}</Text>
                                                    </div>
                                                    <div className={styles.visualhidden}>
                                                        <Label>.</Label>
                                                        <Text className={styles.engagementPrint}>.</Text>
                                                    </div>
                                                </div> : ""}

                                            {this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                                <div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Industry Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.IndustryTypeSelected}</Text>
                                                    </div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Supplement</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.SupplementalSelected}</Text>
                                                    </div>
                                                    <div className={styles.visualhidden}>
                                                        <Label>.</Label>
                                                        <Text className={styles.engagementPrint}>.</Text>
                                                    </div>
                                                </div>
                                                : ""}
                                            {this.state.TeamSelected == 'Advisory' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                                <div>
                                                    <div className={styles.engagementnames}>
                                                        <Label>Template Type</Label>
                                                        <Text className={styles.engagementPrint}>{this.state.AdvisoryTemplateSelected}</Text>
                                                    </div>
                                                    <div className={styles.visualhidden}>
                                                        <Label>.</Label>
                                                        <Text className={styles.engagementPrint}>.</Text>
                                                    </div>
                                                    <div className={styles.visualhidden}>
                                                        <Label>.</Label>
                                                        <Text className={styles.engagementPrint}>.</Text>
                                                    </div>
                                                </div>
                                                : ""}
                                        </div>
                                    </div>
                                    <div className={styles.formcontrols}>
                                        <Label>The following users will automatically have access:</Label>
                                        <div className={styles.usersemail}>{this.state.emailaddress}</div>

                                        {this.state.PortalChoiceSelected == 'Create New' ?
                                            <div>
                                                {this.state.TeamSelected == 'Advisory' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Create New' ?
                                                    <div className={styles.usergroupscopy}>
                                                        {
                                                            this.state.CLUserSelected.split(";").map(element =>
                                                                <div className={styles.usersemails}>{element}</div>
                                                            )
                                                        }
                                                    </div> :
                                                    <div className={styles.userList}>
                                                        {this.state.FinalAccessUserList.split(";").map(element =>
                                                            <div className={styles.usersemail}>{element}</div>
                                                        )}
                                                    </div>
                                                }
                                            </div>
                                            : ""}
                                        {this.state.PortalChoiceSelected == 'Rollover' ?
                                            <div>
                                                {this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow' ?
                                                  <>
                                                    <div className={styles.userLists}>
                                                        <div className={styles.usergroupscopy}>
                                                            <span>CRET-TAX-WF-{this.state.RolloverURL}</span>
                                                            {
                                                                this.state.CRUserSelected.split(";").map(element =>
                                                                    <div className={styles.usersemails}>{element}</div>
                                                                )
                                                            }
                                                        </div>
                                                        <div className={styles.usergroupscopy}>
                                                            <span>CL-TAX-WF-{this.state.RolloverURL}</span>
                                                            {
                                                                this.state.CLUserSelected.split(";").map(element =>
                                                                    <div className={styles.usersemails}>{element}</div>
                                                                )
                                                            }
                                                        </div>
                                                    </div>
                                                    <div>
                                                    <div className={styles.usergroupscopy}>
                                                            {/* This is the lower section of the permissions for adding users who were not part of the rollover:
                                                                NOTE: this should only be CR users as CL users should be added after the portal is created.
                                                            */}
                                                            {this.state.FinalAccessUserList.length != 0 ? <Label>The following new users will automatically have access:</Label> : ""}
                                                            {
                                                                this.state.FinalAccessUserList.split(";").map(element =>
                                                                    <div className={styles.usersemails}>{element}</div>
                                                                )
                                                            }
                                                        </div>
                                                      {/* TODO: testing outputting info to summary screen */}
                                                      { this.state.UsersToRollAlerts.length > 0 &&
                                                        <div className={`${styles.usergroupscopy} ${styles.topdivider}`}>
                                                            <Label>If the following users currently have alerts, they will be created to the new sub-portal:</Label>
                                                            {
                                                                this.state.UsersToRollAlerts.map(element =>
                                                                    <div className={styles.usersemails}>{element.email}</div>
                                                                )
                                                            }
                                                        </div>
                                                        }
                                                    </div>
                                                  </> :
                                                    this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.AssuranceSplitRollover.length == 0 ?
                                                      <>
                                                        <div className={styles.userLists}>
                                                            <div className={styles.usergroupscopy}>
                                                                <span>CRET-AUD-WF-{this.state.RolloverURL}</span>
                                                                {
                                                                    this.state.CRUserSelected.split(";").map(element =>
                                                                        <div className={styles.usersemails}>{element}</div>
                                                                    )
                                                                }
                                                            </div>
                                                            <div className={styles.usergroupscopy}>
                                                                <span>CL-AUD-WF-{this.state.RolloverURL}</span>
                                                                {
                                                                    this.state.CLUserSelected.split(";").map(element =>
                                                                        <div className={styles.usersemails}>{element}</div>
                                                                    )
                                                                }
                                                            </div>
                                                        </div>
                                                      </> : ""}
                                                {this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow' && this.state.PortalChoiceSelected == 'Rollover' ?
                                                    <div>
                                                        {/* Do NOT DELETE THIS CODE */}
                                                        {/* Do NOT DELETE THIS CODE */}
                                                        {/* Do NOT DELETE THIS CODE */}
                                                        {/* <AssuranceEngSplitRolloverDisplayUsers spContext={this.props.spContext} Data={this.state.AssuranceSplitRollover} Control={this.SetAssuranceSplitDataRollOver}></AssuranceEngSplitRolloverDisplayUsers> */}
                                                        <div className={styles.usergroupscopy}>
                                                            {/* This is the lower section of the permissions for adding users who were not part of the rollover:
                                                                NOTE: this should only be CR users as CL users should be added after the portal is created.
                                                            */}
                                                            {this.state.FinalAccessUserList.length != 0 ? <Label>The following new users will automatically have access:</Label> : ""}
                                                            {
                                                                this.state.FinalAccessUserList.split(";").map(element =>
                                                                    <div className={styles.usersemails}>{element}</div>
                                                                )
                                                            }
                                                        </div>
                                                        <div>
                                                          {/* TODO: testing outputting info to summary screen */}
                                                          { this.state.UsersToRollAlerts.length > 0 &&
                                                            <div className={`${styles.usergroupscopy} ${styles.topdivider}`}>
                                                              <Label>If the following users currently have alerts, they will be created in the new sub-portal:</Label>
                                                              {
                                                                this.state.UsersToRollAlerts.map(element =>
                                                                    <div className={styles.usersemails}>{element.email}</div>
                                                                )
                                                              }
                                                            </div>
                                                            }
                                                        </div>
                                                    </div>
                                                    : ""}
                                            </div>
                                            : ""}
                                    </div>
                                    <div className={styles.divider}></div>
                                    <div className={styles.formcontrols}>
                                        <Label>Notifications</Label>
                                        <Checkbox label="Email the above users once subportal has been created." checked={this.state.emailNotification} onChange={(ev, value) => {
                                            this.setState({ emailNotification: value });
                                        }} />
                                    </div>
                                    <div className={styles.formcontrol}>
                                        {this.state.TeamSelected == 'Advisory' &&
                                            <div className={styles.labelprint}>
                                                <DatePicker
                                                    label="Portal Expiration"
                                                    placeholder="MM/DD/YYYY"
                                                    isRequired={true}
                                                    ariaLabel="Select a Portal Expiration Date"
                                                    onSelectDate={this._onSelectDate}
                                                    formatDate={this._onFormatDate}
                                                    minDate={minDate}
                                                    value={advMax} // was this.state.portalExpiration New date set to default to 36 months
                                                    maxDate={advMax}
                                                />
                                                {(this.state.validate && this.state.portalExpiration == null) ?
                                                    <div className={styles.reqval}>Portal Expiration is mandatory.</div> : ''}
                                            </div>
                                        }
                                        {this.state.TeamSelected != 'Advisory' && this.state.PortalTypeSelected == 'Workflow' &&
                                            // Assurance and Tax Workflow portals
                                            <div className={`{styles.labelprint} ${styles.fileExpDatePicker}`}>
                                                <DatePicker
                                                    label="File Expiration"
                                                    placeholder="MM/DD/YYYY"
                                                    isRequired={true}
                                                    ariaLabel="Select a File Expiration Date"
                                                    onSelectDate={this._onSelectDateFileExp}
                                                    formatDate={this._onFormatDate}
                                                    minDate={minDate}
                                                    maxDate={maxDate} // 12 months
                                                    value={maxDate} // now 12 months.  was this.state.DateExtend
                                                />
                                                <div className={styles.fileExpText}>
                                                    Files will be deleted from the portal on this date. The portal will available for rollover for an additional 6 months.
                                                </div>
                                                {(this.state.validate && this.state.DateExtend == null) ?
                                                    <div className={styles.reqval}>File Expiration is mandatory.</div> : ''}
                                            </div>

                                        }
                                        {this.state.TeamSelected != 'Advisory' && this.state.PortalTypeSelected != 'Workflow' &&
                                            // Assurance File Exchange Portals
                                            <div className={styles.labelprint}>
                                                <DatePicker
                                                    label="Portal Expiration"
                                                    placeholder="MM/DD/YYYY"
                                                    isRequired={true}
                                                    ariaLabel="Select a Portal Expiration Date"
                                                    onSelectDate={this._onSelectDate2}
                                                    formatDate={this._onFormatDate}
                                                    minDate={minDate}
                                                    maxDate={maxDate} // 12 months
                                                    value={maxDate} // now 12 months.  was this.state.DateExtend
                                                />
                                                {(this.state.validate && this.state.portalExpiration == null) ?
                                                    <div className={styles.reqval}>Portal Expiration is mandatory.</div> : ''}
                                            </div>
                                        }
                                    </div>
                                    {this.state.success && this.state.IsPortalEntryCreated == "Y" ?
                                        <div>
                                            <Label className={styles.successMsg}>Thank you. Your portal is in the process of being created. You will receive an email notification shortly when your portal is active. Please close this window.</Label>
                                        </div> : ""
                                    }
                                    {!this.state.success && this.state.IsPortalEntryCreated == "N" ?
                                        <div>
                                            <Label className={styles.errormsg}>Something went wrong. Please refresh page and try to submit request again</Label>
                                        </div> : ""
                                    }
                                </div> : ""}

                            {this.state.currentScreen == 'screen5' ?
                                <div className={styles.screenFour}>
                                    <div className={styles.freshRollover}>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Portal Type</Label>
                                                <Text className={styles.engagementPrint}>{this.state.PortalTypeSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Number</Label>
                                                <Text className={styles.engagementPrint}>
                                                    {updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Engagement Name</Label>
                                                <Text className={styles.engagementPrint}>{this.state.EngagementNameSelected}</Text>
                                            </div>
                                        </div>
                                        <div className={styles.engnumbername}>
                                            <div className={styles.engagementnames}>
                                                <Label>Work Year</Label>
                                                <Text className={styles.engagementPrint}>{this.state.Year}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>Site Owner</Label>
                                                <Text className={styles.engagementPrint}>{this.state.addusers}</Text>
                                            </div>
                                            <div className={styles.engagementnames}>
                                                <Label>.</Label>
                                                <Text className={styles.engagementPrint}>.</Text>
                                            </div>
                                        </div>
                                    </div>
                                    <div className={styles.formcontrols}>
                                        <Label>Upload Investors</Label>
                                        <input type="file" onChange={this.OnFileSelect} id='newfile' name='newfile'
                                        ></input>
                                        <label className={styles.browsebutton} htmlFor={"newfile"}><span>Choose File</span></label>
                                        {this.state.K1FileName != "" ? <p className={styles.addedFile}><b>{this.state.K1FileName}</b> is selected to upload </p> : null}
                                        {(this.state.validate && this.state.K1FileName == "") ?
                                            <div className={styles.reqval}>Upload Investors is mandatory.</div> : ''}
                                        <p><i>Please upload a properly formatted Excel document containing all investor names and their corresponding access. You may download an Excel template for this process here.</i></p>
                                        <a className={styles.downloadEXL} id="download_link" download={HubSiteURL + K1ExcelTemplate} href={HubSiteURL + K1ExcelTemplate}><Icon iconName="Download" className={styles.Icon} />Download Excel Template</a>
                                        {/* Do NOT DELETE THIS CODE */}
                                        {/* Do NOT DELETE THIS CODE */}
                                        {/* Do NOT DELETE THIS CODE */}
                                        {/* <a className={styles.k1instruction} id="K1-instruction" target="_blank" href={HubSiteURL + '/SitePages/K1-Portal%20Instructions.aspx'}>K1-Portal Instructions</a> */}
                                    </div>
                                    <div className={styles.divider}></div>
                                    <div className={styles.formcontrols}>
                                        <Label>Notifications</Label>
                                        <Checkbox label="Email the above users once subportal has been created." checked={this.state.emailNotification} onChange={(ev, value) => {
                                            this.setState({ emailNotification: value });
                                        }} />
                                    </div>
                                    <div className={styles.formcontrol}>
                                        <div className={styles.labelprint}>
                                            <DatePicker
                                                label="Portal Expiration"
                                                placeholder="MM/DD/YYYY"
                                                isRequired={false}
                                                ariaLabel="Select a Portal Expiration Date"
                                                onSelectDate={this._onSelectDate3}
                                                formatDate={this._onFormatDate}
                                                minDate={minDate}
                                                maxDate={K1MaxDate}
                                                value={this.state.K1Date}
                                            />
                                        </div>
                                    </div>
                                    {this.state.success && this.state.IsPortalEntryCreated == "Y" ?
                                        <div>
                                            <Label className={styles.successMsg}>Thank you. Your portal is in the process of being created. You will receive an email notification shortly when your portal is active. Please close this window.</Label>
                                        </div> : ""
                                    }
                                    {!this.state.success && this.state.IsPortalEntryCreated == "N" ?
                                        <div>
                                            <Label className={styles.errormsg}>Something went wrong. Please refresh page and try to submit request again</Label>
                                        </div> : ""
                                    }
                                </div> : ""}
                        </div>
                        {(this.state.IsDuplicate) ?
                            <div className={styles.errormsg}>A duplicate portal with same Portal ID already exists. Please try to create a new one..</div> : ''}

                        <DialogFooter className={styles.buttonFooter}>
                            <DefaultButton text={this.state.cancelbuttonname} onClick={this.CloseButton} />
                            <div>
                                {this.state.currentScreen != "screen1" ?
                                    <DefaultButton className={styles.backButton} disabled={this.state.disableBtn} text="Back" onClick={(event) => {
                                        this.previousDialog(event);
                                    }} /> : ""}
                                <PrimaryButton
                                    className={styles.NextButton}
                                    disabled={this.state.disableBtn}
                                    text={this.state.dialogbuttonname}
                                    onClick={(event) => {
                                        this.submitDialog(event);
                                    }}
                                />
                            </div>
                        </DialogFooter>
                    </div>
                </Dialog>
            </div >
        );
    }

    private previousDialog(e) {
        if (this.state.currentScreen == "screen2" || this.state.currentScreen == "screen5") {
            this.setState({
                validate: false,
                IsDuplicate: false,
                currentScreen: "screen1",
                dialogbuttonname: "Next",
                titleText: "",

            });
        }
        else if (this.state.currentScreen == "screen3") {
            this.setState({
                validate: false,
                IsDuplicate: false,
                currentScreen: "screen2",
                dialogbuttonname: "Next",
                titleText: "- Rollover - Create New",

            });
        }
        else if (this.state.currentScreen == "screen4") {

            this.setState({
                validate: false,
                IsDuplicate: false,
                currentScreen: "screen3",
                dialogbuttonname: "Next",
                titleText: "- Template and Provisioning",
            });
        }
    }

    private submitDialog = async (e) => {
      console.log('logging this.state.AdvisoryTemplate', this.state.AdvisoryTemplate);

        if (this.state.currentScreen == "screen1") {
            if (this.state.EngagementNumberSelected == "" || this.state.addusers.length == 0
                || this.state.Year == null || this.state.PortalTypeSelected == ""
                || (this.state.TeamSelected == "" && this.state.PortalTypeSelected == 'K1')
            ) {
                this.setState({
                    validate: true
                });

            } else {
                this.checkEngagement(this.state.PortalsCreated);
                this.newEngagementNumber();
                if (this.state.PortalTypeSelected == 'K1' && this.state.Checkeng == true) {
                    this.setState({
                        validate: false,
                        currentScreen: "screen5",
                        dialogbuttonname: "Create Portal",
                        titleText: "- Confirmation",
                    });
                }
                else {
                    if (this.state.Checkeng == true) {
                        if (this.state.TeamSelected == 'Tax' && this.state.PortalTypeSelected == 'Workflow') {
                            this.Rollover();
                        }
                        else if (this.state.TeamSelected == 'Assurance' && this.state.PortalTypeSelected == 'Workflow') {
                            this.Rollover();
                            //this.CheckSplitRollover();
                        }
                        else {
                            this.setState({ isRollover: true });
                        }
                        this.setState({
                            validate: false,
                            currentScreen: "screen2",
                            dialogbuttonname: "Next",
                            titleText: "- Rollover - Create New",
                        });
                    }
                }
            }
        } else if (this.state.currentScreen == "screen2") {
            if (this.state.addusers.length == 0 || this.state.PortalChoiceSelected == "") {
                this.setState({
                    validate: true
                });
            } else {
                if (!this.state.IndustryType.length) {
                  this.loadIndustryTypes();
                }
                if (!this.state.ServiceType.length) {
                  this.loadServiceTypes();
                }
                if (!this.state.AdvisoryTemplate.length) {
                  this.loadAdvisoryTemplates();
                }
                this.setState({
                    validate: false,
                    currentScreen: "screen3",
                    dialogbuttonname: "Next",
                    titleText: "- Template and Provisioning",
                });
            }
        } else if (this.state.currentScreen == "screen3") {
            if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Advisory") {

                if (this.state.addusers.length == 0 || this.state.AdvisoryTemplateSelected == "") {
                    this.setState({
                        validate: true
                    });
                }
                else {
                    if (this.state.validate != true) {
                        this.setState({
                            validate: false,
                            currentScreen: "screen4",
                            dialogbuttonname: "Create Portal",
                            titleText: "- Confirmation",
                        });
                        this.getCLUserList();
                    }
                }
            }
            else if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Assurance") {

                if (this.state.PortalChoiceSelected == "Create New") {
                    if (this.state.addusers.length == 0 || this.state.IndustryTypeSelected == "" || this.state.SupplementalSelected == "") {
                        this.setState({
                            validate: true
                        });
                    } else {
                        if (this.state.validate != true) {
                            let txtval = true;
                            this.state.AsuranceSplitData.txtValues.map(ev => {
                                if (ev.Value == "") {
                                    txtval = false;
                                }
                            });
                            if (this.state.AsuranceSplitData.disabled == false && txtval != true) {
                                this.setState({
                                    validate: true
                                });
                                let ErrorMessage = "Please enter correct Split data";
                                this.setState({ Message: ErrorMessage, showMessageBar: true, MessageBarType: OfficeUI.MessageBarType.error, disableBtn: false });
                            } else {
                                this.setState({
                                    validate: false,
                                    currentScreen: "screen4",
                                    dialogbuttonname: "Create Portal",
                                    titleText: "- Confirmation",
                                });
                                this.getCLUserList();
                            }
                        }
                    }
                }
                else if (this.state.PortalChoiceSelected == "Rollover") {
                    let txtval = true;

                    this.state.AssuranceSplitRollover.map(ev => {
                        if (ev.NewSplitValue == "") {
                            txtval = false;
                        }
                    });

                    if (txtval != true) {
                        this.setState({
                            validate: true
                        });
                        let ErrorMessage = "Please enter correct Split data";
                        this.setState({ Message: ErrorMessage });
                        this.setState({ showMessageBar: true });
                        this.setState({ MessageBarType: OfficeUI.MessageBarType.error });
                        this.setState({ disableBtn: false });
                    } else {
                        this.setState({
                            validate: false,
                            currentScreen: "screen4",
                            dialogbuttonname: "Create Portal",
                            titleText: "- Confirmation",
                        });
                    }
                    if (this.state.AssuranceSplitRollover.length == 0) {
                        this.getCLUserList();

                    }
                }

            }
            else if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Tax") {
                if (this.state.addusers.length == 0 || this.state.IndustryTypeSelected == "" || this.state.ServiceTypeSelected == "") {
                    this.setState({
                        validate: true
                    });
                }
                else {
                    this.setState({
                        validate: false,
                        currentScreen: "screen4",
                        dialogbuttonname: "Create Portal",
                        titleText: "- Confirmation",
                    });
                    this.getCLUserList();
                }
            }

            else {
                if (this.state.addusers.length == 0) {
                    this.setState({
                        validate: true
                    });
                }
                else {
                    if (this.state.validate != true) {
                        this.setState({
                            validate: false,
                            currentScreen: "screen4",
                            dialogbuttonname: "Create Portal",
                            titleText: "- Confirmation",
                        });
                        this.getCLUserList();
                    }
                }
            }
        } else {
            let PortalId = "";
            let FinalEngNumber = updatedworkyear == true ? this.state.UpdatedEngagementNumberSelected : this.state.EngagementNumberSelected;

            if (this.state.PortalTypeSelected == "Workflow" && this.state.TeamSelected == "Advisory") {
                if (this.state.portalExpiration == null) {
                    this.setState({
                        validate: true
                    });
                }
                else {
                    PortalId = this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
                    let _isDuplicatePortal = await this.CheckDuplicateAdvantagePortal(PortalId);
                    if (!_isDuplicatePortal) {
                        this.save();
                    }
                    else {
                        this.setState({
                            IsDuplicate: true,
                        });
                    }
                }
            }
            else if (this.state.PortalTypeSelected == 'K1') {
                if (this.state.K1FileName == "") {
                    this.setState({
                        validate: true
                    });
                }
                else {
                    PortalId = "TAX-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
                    let _isDuplicatePortal = await this.CheckDuplicateAdvantagePortal(PortalId);
                    if (!_isDuplicatePortal) {
                        this.save();
                    }
                    else {
                        this.setState({
                            IsDuplicate: true,
                        });
                    }
                }
            }
            else {
                    PortalId = this.state.TeamURL + "-" + this.state.PortalTypeURL + "-" + FinalEngNumber;
                    let _isDuplicatePortal = await this.CheckDuplicateAdvantagePortal(PortalId);
                    if (!_isDuplicatePortal) {
                        this.save();
                    }
                    else {
                        this.setState({
                            IsDuplicate: true,
                        });
                    }
            }
        }
    }

    private save() {
        let siteCollectionUrl = GlobalValues.HubSiteURL;
        let listname = GlobalValues.EngagementPortalList;
        this.SaveItem(siteCollectionUrl, listname)
            .then(() => {
                this.ShowHideProgressBar(false);
                if (this.state.IsPortalEntryCreated = "Y") {
                    this.setState({
                        success: true,
                        disableBtn: true,
                        cancelbuttonname: "Close",
                    });
                }
                else {
                    this.setState({
                        success: false,
                        disableBtn: true,
                        cancelbuttonname: "Close",
                    });
                }
            });
    }

    private CheckDuplicateAdvantagePortal = (async (_currPortalId) => {
        let _isDuplicate: boolean = false;
        const caml: ICamlQuery = {
            ViewXml: "<View><Query><Where><Eq><FieldRef Name='PortalId'/><Value Type='Text'>" + _currPortalId + "</Value></Eq></Where></Query></View>",
        };
        let hubWeb = Web(GlobalValues.HubSiteURL);
        return await hubWeb.lists.getByTitle("Engagement Portal List").getItemsByCAMLQuery(caml).then((data) => {
            return data.length > 0 ? _isDuplicate = true : _isDuplicate = false;
        });
    });

}

export default CreateEngagement;
