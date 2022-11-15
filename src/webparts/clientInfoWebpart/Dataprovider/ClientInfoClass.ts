import { GlobalValues } from "./GlobalValue";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import { ICamlQuery } from "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";

export class ClientInfoClass {
  public _SetupSP() {
    sp.setup({
      sp: {
        baseUrl: GlobalValues.SiteURL,
      },
      ie11: true,
    });
    return sp;
  }

  public HubSite_SetupSP() {
    sp.setup({
      sp: {
        baseUrl: GlobalValues.HubSiteURL,
      },
      ie11: true,
    });
    return sp;

  }

  public GetClientInfo = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    this.HubSite_SetupSP();

    let ClientInformationData = {
      LinkTitle: "",
      ClientNumber: "",
    };

    let absoluteUrl = GlobalValues.SiteURL;
    let finalabsoluteUrl = absoluteUrl.split("/");
    let CRN = finalabsoluteUrl[finalabsoluteUrl.length - 1];
    await sp.web.lists
      .getByTitle(GlobalValues.ClientList)
      .items.filter("ClientNumber eq '" + CRN + "'").getAll()
      .then((results) => {
        let data = {
          ClientNumber: "",
          Title: ""
        };
        let browserUrl = window.location.href;
        let ClientNum = browserUrl.split("/")[4];

        for (let i = 0; i < results.length; i++) {
          let ClientNumber = results[i].ClientNumber;
          if (ClientNumber == ClientNum) {
            data = results[i];
            break;
          }
        }
        ClientInformationData.ClientNumber = data.ClientNumber;
        ClientInformationData.LinkTitle = data.Title;
      });

    return ClientInformationData;
  }

  public GetEngInfo = async () => {
    let absoluteUrl = GlobalValues.SiteURL;
    let finalabsoluteUrl = absoluteUrl.split("/");
    let CRN = finalabsoluteUrl[finalabsoluteUrl.length - 1];
    this.HubSite_SetupSP();

    const Engagementdata = await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.EngagementList)
      .items.filter("ClientNumber eq '" + CRN + "'").getAll()
      .then((results) => {
        results = results.filter(e => !e.PortalExists && (e.PortalProgress != "In Progress" && e.PortalProgress != "Completed"));
        return results;
      });
    return Engagementdata;
  }

  public GetEngPortalListItemID = async (_currPortalId) => {
    this.HubSite_SetupSP();
    const caml: ICamlQuery = {
      ViewXml: "<View><Query><Where><Eq><FieldRef Name='PortalId'/><Value Type='Text'>" + _currPortalId + "</Value></Eq></Where></Query></View>",
    };
    return await sp.web.lists.getByTitle(GlobalValues.EngagementPortalList).getItemsByCAMLQuery(caml).then((data) => {
      return data[0].ID;
    });
  }

  public GetAdvisoryTemplates = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));

    this.HubSite_SetupSP();

    const AdvisoryTemplatesdata = await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.AdvisoryTemplatesList)
      .items.orderBy('Title', true).get()
      .then((results) => {
        return results;
      });
    return AdvisoryTemplatesdata;
  }


  public GetIndustryTypes = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    this.HubSite_SetupSP();


    const IndustryTypesdata = await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.IndustryTypesList)
      .items.orderBy('Title', true).getAll()
      .then((results) => {
        return results;
      });
    return IndustryTypesdata;
  }

  public GetSupplemental = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    this.HubSite_SetupSP();

    const Supplementaldata = await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.AssuranceSupplementalList)
      .items.orderBy('Title', true).getAll()
      .then((results) => {
        return results;
      });
    return Supplementaldata;
  }

  public GetServiceTypes = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    this.HubSite_SetupSP();


    const ServiceTypesdata = await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.ServiceTypesList)
      .items.orderBy('Title', true).getAll()
      .then((results) => {
        return results;
      });
    return ServiceTypesdata;
  }


  public GetUsersByGroup = (async (groupName) => {
    this._SetupSP();
    const users = await this._SetupSP().web.siteGroups.getByName(groupName).users();
    return users;
  });

  public SaveEngagementList = async (PortalsCreated, id) => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    this.HubSite_SetupSP();
    let objToSave = {
      "Portals_x0020_Created": PortalsCreated,
    };

    await this.HubSite_SetupSP().web.lists.getByTitle(GlobalValues.EngagementList).items.getById(id).update(objToSave).then((data) => {
      return true;
    });
  }


  public CheckIfEngCreated = (async (EngNo) => {
    let query = "";
    query = "Title eq '" + EngNo + "'";

    return await this.HubSite_SetupSP().web.lists
      .getByTitle(GlobalValues.EngagementPortalList)
      .items.filter(query).getAll()
      .then((results) => {
        if (results.length > 0) {
          return true;
        }
        else {
          return false;
        }

      });

  });
}
