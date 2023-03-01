
import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import { GlobalValues } from "./GlobalValue";
import { ClientInfoClass } from "../Dataprovider/ClientInfoClass";

export class AssuranceSplitRolloverModel {
  public TxtId;
  public ToggleId;
  public RecordId;
  public RollOverEngNumber;
  public CurrentEngNumber;
  public OldEngURL;
  public NewEngURL;
  public OldSplitValue;
  public NewSplitValue;
  public SplitRollOverURL;
  public IsSplitEngRollOver;
  public IsNewRecord;
  public CRETGroupName;
  public CLGroupName;
  public CRETGroupUsers;
  public CLGroupUsers;
  public NewPortalId;
  public CreateasRollover;
  public Errormessage;
  public SplitCategory;

}

export class AssuranceSplitRollover {
/*
  public HubSite_SetupSP() {
    sp.setup({
      sp: {
        baseUrl: GlobalValues.HubSiteURL,
      }
    });
    return sp;

  }
*/
  public uuidv4() {
    return GlobalValues.uuidv4String.replace(/[xy]/g, (c) => {
      var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  public SaveSplitEngagementRollover = async (RequestObj, RequestData) => {

    let objReq: any;
    let promiseArr = [];
    RequestData.forEach(element => {

      objReq = { ...RequestObj };
      if (element.CreateasRollover == true) {

        if (element.IsNewRecord == false) {
          let CLuserList = "";
          element.CLGroupUsers.filter(ev => ev.checked == true).forEach((e) => {
            CLuserList += e.email + ";";
          });
          let CRuserList = "";
          element.CRETGroupUsers.filter(ev => ev.checked == true).forEach((e) => {
            CRuserList += e.email + ";";
          });

          let PortalData: any = {
            'EngagementName': objReq.EngagementName + "-" + element.NewSplitValue,
            'Title': objReq.Title,
            'EngagementNumberEndZero': objReq.EngagementNumberEndZero,
            'ClientNumber': objReq.ClientNumber,
            'PortalType': objReq.PortalType,
            'Team': objReq.Team,
            'WorkYear': objReq.WorkYear,
            'SiteOwnerId': objReq.SiteOwnerId,
            'SiteUrl': { Url: objReq.SiteUrl.Url + "-" + element.NewSplitValue },
            'EngagementMembers': CRuserList,
            'ClientMembers': CLuserList + objReq.ClientMembers,
            'Rollover': objReq.Rollover,
            'RolloverUrl': { Url: element.OldEngURL },
            'IndustryType': objReq.IndustryType,
            'Supplemental': objReq.Supplemental,
            'TemplateType': objReq.TemplateType,
            'isNotificationEmail': objReq.isNotificationEmail,
            'PortalExpiration': objReq.PortalExpiration,
            'PortalId': objReq.PortalId + "-" + element.NewSplitValue,
            'SplitEngNumber': element.CurrentEngNumber,
            'SplitCategory': element.SplitCategory,
            'SplitSuffix': element.NewSplitValue,
            'WorkpaperPath': objReq.WorkpaperPath
          };

          promiseArr.push(new Promise((resolve, reject) => {
            return resolve(sp.web.lists.getByTitle(GlobalValues.EngagementPortalList).items.add(PortalData));
          }));
        }

        else if (element.IsNewRecord == true) {

          let PortalData: any = {
            'EngagementName': objReq.EngagementName + "-" + element.NewSplitValue,
            'Title': objReq.Title,
            'EngagementNumberEndZero': objReq.EngagementNumberEndZero,
            'ClientNumber': objReq.ClientNumber,
            'PortalType': objReq.PortalType,
            'Team': objReq.Team,
            'WorkYear': objReq.WorkYear,
            'SiteOwnerId': objReq.SiteOwnerId,
            'SiteUrl': { Url: objReq.SiteUrl.Url + "-" + element.NewSplitValue },
            'EngagementMembers': "",
            'ClientMembers': objReq.ClientMembers,
            'Rollover': false,
            'IndustryType': objReq.IndustryType,
            'Supplemental': objReq.Supplemental,
            'TemplateType': objReq.TemplateType,
            'isNotificationEmail': objReq.isNotificationEmail,
            'PortalExpiration': objReq.PortalExpiration,
            'PortalId': objReq.PortalId + "-" + element.NewSplitValue,
            'SplitEngNumber': objReq.Title,
            'SplitCategory': "",
            'SplitSuffix': element.NewSplitValue,
            'WorkpaperPath': objReq.WorkpaperPath
          };

          promiseArr.push(new Promise((resolve, reject) => {
            return resolve(sp.web.lists.getByTitle(GlobalValues.EngagementPortalList).items.add(PortalData));
          }));
        }
      }
    });

    // In parallel with Promise.all
    return await Promise.all(promiseArr).then(values => {
      return true;
    }).catch(error => {
      console.log("SaveSplitEngagementRollover::Error:", error);
      return false;
    });

  }


  public GetAssuranceSplitRollover = async (EngNo, CurrentEngNo, Teamtype, PortalType, Isnextyear) => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    //this.HubSite_SetupSP();
    let counter = 1;
    let query = "";
    if (Isnextyear == true) {
      query = "Title eq '" + EngNo + "'";
    }
    else {
      query = "Title eq '" + EngNo + "'";
    }

    let hubWeb = Web(GlobalValues.ClientPortalURL);
    return await hubWeb.lists
      .getByTitle(GlobalValues.EngagementPortalList)
      .items.filter(query).getAll()
      .then((results) => {
        let RollOverSplitData = [];

        results.forEach((element) => {
          let obj = new AssuranceSplitRolloverModel();
          let NewPortalId = "";
          if (Teamtype == "Assurance" && PortalType == "Workflow") {
            NewPortalId = "AUD-WF-" + CurrentEngNo;
          }
          else if (Teamtype == "Assurance" && PortalType == "File Exchange") {
            NewPortalId = "AUD-FE-" + CurrentEngNo;
          }

          obj.RollOverEngNumber = EngNo;
          obj.CurrentEngNumber = CurrentEngNo;
          obj.OldEngURL = element.SiteUrl.Url;
          obj.RecordId = this.uuidv4();
          obj.OldSplitValue = element.SplitSuffix;
          obj.NewSplitValue = element.SplitSuffix;
          obj.SplitRollOverURL = element.SiteUrl.Url;
          obj.IsSplitEngRollOver = true;
          obj.CreateasRollover = true;
          obj.IsNewRecord = false;
          obj.CRETGroupName = "CRET" + "-" + element.PortalId;
          obj.CLGroupName = "CL" + "-" + element.PortalId;
          let UserGroupObj = new ClientInfoClass();
          obj.CRETGroupUsers = [];
          UserGroupObj.GetUsersByGroup(obj.CRETGroupName).then((res) => {
            res.forEach((e) => {
              obj.CRETGroupUsers.push({ email: e.Email, checked: true });
            });

          });

          obj.CLGroupUsers = [];
          UserGroupObj.GetUsersByGroup(obj.CLGroupName).then((res) => {
            res.forEach((e) => {
              obj.CLGroupUsers.push({ email: e.Email, checked: true });
            });

          });
          obj.NewPortalId = NewPortalId + "-" + element.SplitSuffix;
          obj.TxtId = "txt" + counter;
          obj.ToggleId = "tog" + counter;
          counter = counter + 1;
          obj.SplitCategory = element.SplitCategory;
          RollOverSplitData.push(obj);
        });
        if (RollOverSplitData.length == 1 && RollOverSplitData[0].OldSplitValue == null && RollOverSplitData[0].SplitCategory == null) {
          return [];
        }
        else {
          return RollOverSplitData;
        }
      });
  }
}