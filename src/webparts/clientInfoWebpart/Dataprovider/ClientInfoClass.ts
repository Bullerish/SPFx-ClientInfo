import { GlobalValues } from "./GlobalValue";
import { sp } from "@pnp/sp";
import { ICamlQuery } from "@pnp/sp/lists";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";

export class ClientInfoClass {

  public GetClientInfo = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
      let ClientInformationData = {
      LinkTitle: "",
      ClientNumber: "",
    };

    let absoluteUrl = GlobalValues.SiteURL;
    let finalabsoluteUrl = absoluteUrl.split("/");
    let CRN = finalabsoluteUrl[finalabsoluteUrl.length - 1];



    const hubWeb = Web(GlobalValues.HubSiteURL);
    await hubWeb.lists
      .getByTitle(GlobalValues.ClientList)
      .items.filter("ClientNumber eq '" + CRN + "'").getAll()
      .then((results) => {
        ClientInformationData.ClientNumber = results[0].ClientNumber;
        ClientInformationData.LinkTitle = results[0].Title;
      });

    return ClientInformationData;
  }

  public GetEngInfo = async () => {
    let absoluteUrl = GlobalValues.SiteURL;
    let finalabsoluteUrl = absoluteUrl.split("/");
    let CRN = finalabsoluteUrl[finalabsoluteUrl.length - 1];
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const Engagementdata = await hubWeb.lists
      .getByTitle(GlobalValues.EngagementList)
      .items.filter("ClientNumber eq '" + CRN + "'").getAll()
      .then((results) => {
        results = results.filter(e => !e.PortalExists && (e.PortalProgress != "In Progress" && e.PortalProgress != "Completed"));
        return results;
      });
    console.log('get eng info',Engagementdata);
    return Engagementdata;
  }

  public GetEngPortalListItemID = async (_currPortalId) => {
    const caml: ICamlQuery = {
      ViewXml: "<View><Query><Where><Eq><FieldRef Name='PortalId'/><Value Type='Text'>" + _currPortalId + "</Value></Eq></Where></Query></View>",
    };
    const hubWeb = Web(GlobalValues.HubSiteURL);
    return await hubWeb.lists.getByTitle(GlobalValues.EngagementPortalList).getItemsByCAMLQuery(caml).then((data) => {
      return data[0].ID;
    });
  }

  public GetAdvisoryTemplates = async () => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const AdvisoryTemplatesdata = await hubWeb.lists
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
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const IndustryTypesdata = await hubWeb.lists
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
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const Supplementaldata = await hubWeb.lists
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
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const ServiceTypesdata = await hubWeb.lists
      .getByTitle(GlobalValues.ServiceTypesList)
      .items.orderBy('Title', true).getAll()
      .then((results) => {
        return results;
      });
    return ServiceTypesdata;
  }


  public GetUsersByGroup = (async (groupName) => {
    const clientWeb = Web(GlobalValues.SiteURL);
    const users = await clientWeb.siteGroups.getByName(groupName).users();
    return users;
  });

  public SaveEngagementList = async (PortalsCreated, id) => {
    let url = GlobalValues.SiteURL;
    url = url.substring(0, url.lastIndexOf("/"));
    let objToSave = {
      "Portals_x0020_Created": PortalsCreated,
    };
    const hubWeb = Web(GlobalValues.HubSiteURL);
    await hubWeb.lists.getByTitle(GlobalValues.EngagementList).items.getById(id).update(objToSave).then((data) => {
      return true;
    });
  }


  public CheckIfEngCreated = (async (EngNo) => {
    let query = "";
    query = "Title eq '" + EngNo + "'";
    const hubWeb = Web(GlobalValues.HubSiteURL);
    return await hubWeb.lists
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
