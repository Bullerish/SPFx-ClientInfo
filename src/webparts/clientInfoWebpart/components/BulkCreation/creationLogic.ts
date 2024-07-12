import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { GlobalValues } from "../../Dataprovider/GlobalValue";

const hubSite: IWeb = Web(GlobalValues.HubSiteURL);

export const getMatterNumbersForClientSite = async (clientSiteNumber: string): Promise<any[]> => {
  const engagementListMatters = await hubSite.lists.getByTitle("Engagement List").items.filter(`ClientNumber eq '${clientSiteNumber}' and Team ne 'ADV'`).get();
  return engagementListMatters;
};

export const createEngagementSubportals = async (selectedEngagements: any[], portalType: string, team: string) => {
  for (const engagement of selectedEngagements) {
    const matterNumberYear = engagement.Title.split('-')[1]; // Assuming the year is part of the matter number
    const year = matterNumberYear ? matterNumberYear : ''; // Default to an empty string if year is not found

    const data = {
      Title: engagement.Title,
      ClientNumber: engagement.ClientNumber,
      EngagementName: engagement.EngagementName,
      WorkYear: year,
      Team: team,
      PortalType: portalType,
      IndustryType: engagement.IndustryType,
      Supplemental: engagement.Supplemental,
      SiteOwner: engagement.SiteOwner,
      ExpirationDate: portalType === 'fileExchange' ? createFileExpirationDate() : createDate18MonthsFromNow(),
    };

    await hubSite.lists.getByTitle("Engagement Portal List").items.add(data);
  }
};

const createDate18MonthsFromNow = (): Date => {
  const date = new Date();
  date.setMonth(date.getMonth() + 18);
  return date;
};

const createFileExpirationDate = (): Date => {
  const date = new Date();
  date.setMonth(date.getMonth() + 12);
  return date;
};
