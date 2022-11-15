import { sp } from "@pnp/sp";
import { GlobalValues } from "./GlobalValue";
export class AssuranceSplit {
    public SaveSplitEngagement = async (RequestObj, RequestData) => {

        let promiseArr = [];
        let objReq: any;
        RequestData.txtValues.forEach(element => {
            objReq = { ...RequestObj };

            let PortalData: any = {

                'EngagementName': objReq.EngagementName + "-" + element.Value,
                'Title': objReq.Title,
                'EngagementNumberEndZero': objReq.EngagementNumberEndZero,
                'ClientNumber': objReq.ClientNumber,
                'PortalType': objReq.PortalType,
                'Team': objReq.Team,
                'WorkYear': objReq.WorkYear,
                'SiteOwnerId': objReq.SiteOwnerId,
                'SiteUrl': { Url: objReq.SiteUrl.Url + "-" + element.Value },
                'EngagementMembers': objReq.EngagementMembers,
                'ClientMembers': objReq.ClientMembers,
                'Rollover': objReq.Rollover,
                'IndustryType': objReq.IndustryType,
                'Supplemental': objReq.Supplemental,
                'TemplateType': objReq.TemplateType,
                'isNotificationEmail': objReq.isNotificationEmail,
                'PortalExpiration': objReq.PortalExpiration,
                'PortalId': objReq.PortalId + "-" + element.Value,
                'SplitEngNumber': objReq.Title,
                'SplitCategory': RequestData.SelectedCategory,
                'SplitSuffix': element.Value,
                'WorkpaperPath': objReq.WorkpaperPath

            };

            promiseArr.push(new Promise((resolve, reject) => {
                return resolve(sp.web.lists.getByTitle(GlobalValues.EngagementPortalList).items.add(PortalData));
            }));
        });

        // In parallel with Promise.all
        return await Promise.all(promiseArr).then(values => {
            return true;
        }).catch(error => {
            console.log("SaveSplitEngagement:: promise arr", error);
            return false;
        });

    }
}