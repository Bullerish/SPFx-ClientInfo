import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
export class GlobalValues {
  public static SiteURL = "";
  public static PermissionPage = "/SitePages/ManageClientPermissions.aspx";

  public static isCRADUser = false;
  public static isCRETUser = false;
  public static isCLUser = false;
  public static EngagementPortalList = "Engagement Portal List";
  public static ClientList = "Client";
  public static EngagementList = "Engagement List";
  public static AdvisoryTemplatesList = "AdvisoryTemplates";
  public static IndustryTypesList = "IndustryTypes";
  public static AssuranceSupplementalList = "Assurance Supplemental";
  public static ServiceTypesList = "ServiceTypes";
  public static K1InvestorDocumentsList = "K1 Investor Documents";
  public static K1InvestorDocumentsURL = "/sites/ClientPortal/K1InvestorDocuments/";
  public static serverRelativeUrl = "";
  public static HubSiteURL = window.location.origin + '/sites/ClientPortal';
  public static uuidv4String = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
  public static errorTitle = "Error";
  public static errorMsg = "An error occurred. Please try again.";

  public static SetValues = (async (context: any) => {
    if (GlobalValues.SiteURL == "")
      GlobalValues.SiteURL = context.pageContext.web.absoluteUrl;
    if (GlobalValues.serverRelativeUrl == "")
      GlobalValues.serverRelativeUrl =
        context.pageContext.web.serverRelativeUrl;
    let alertWeb = Web(GlobalValues.SiteURL);
    return await alertWeb.currentUser.groups().then((usergroups) => {
      if (usergroups.filter(x => x.Title.indexOf("CRAD-AT") > -1
        || x.Title.indexOf("CRAD-ADV") > -1).length > 0) {
        GlobalValues.isCRADUser = true;
      }
      else if (
        usergroups.filter(x => x.Title.indexOf("CRET-TAX") > -1
          || x.Title.indexOf("CRET-ADV") > -1 || x.Title.indexOf("CRET-AUD") > -1).length > 0) {
        GlobalValues.isCRETUser = true;
      }
      else if (
        usergroups.filter(x => x.Title.indexOf("CL-TAX") > -1
          || x.Title.indexOf("CL-ADV") > -1 || x.Title.indexOf("CL-AUD") > -1).length > 0) {
        GlobalValues.isCLUser = true;
      }
      return true;

    });

  });

}
