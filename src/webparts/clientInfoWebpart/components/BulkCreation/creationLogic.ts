import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

const hubSite = Web(GlobalValues.HubSiteURL);

export interface MatterAndCreationData {
  Title?: string;
  ID: string;
  engListID: string;
  newMatterNumber: string;
  creationMatterNumber: string;
  newMatterId: string;
  newMatterSiteUrl: string;
  creationMatterSiteUrl: string;
  templateType: string;
  newMatterPortalId: string;
  newMatterEngagementName: string;
  creation: boolean;
  clientNumber: string;
  team: string;
  newMatterWorkYear: string;
  creationMatterWorkYear: string;
  portalType: string;
  engagementNumberEndZero?: string;
  industryType?: string;
  supplemental: string;
  newMatterPortalExpirationDate: string;
  newMatterFileExpirationDate: string;
  isNotificationEmail: boolean;
  siteOwner: string | ISiteUserInfo | number;
  Portals_x0020_Created?: string;
}

export const createDate18MonthsFromNow = (): Date => {
  const date = new Date();
  date.setMonth(date.getMonth() + 18);
  return date;
};

const createFileExpirationDate = (): Date => {
  const date = new Date();
  date.setMonth(date.getMonth() + 12);
  return date;
};

export const getMatterNumbersForClientSite = async (
  clientSiteNumber: string
): Promise<{
  engagementListMatters: MatterAndCreationData[];
}> => {
  let engagementListMatters: MatterAndCreationData[] = [];

  console.log("getMatterNumbersForClientSite firing::", clientSiteNumber);

  const batchSize = 5000;
  let items: any[] = [];
  let currentPageItems: any[];

  do {
    currentPageItems = await hubSite.lists.getByTitle("Engagement List")
      .items.filter(`ClientNumber eq '${clientSiteNumber}'`)
      .select("Title", "ClientNumber", "EngagementName", "ID", "WorkYear", "Team", "Portals_x0020_Created")
      .top(batchSize)
      .skip(items.length)
      .get();

    items = items.concat(currentPageItems);
  } while (currentPageItems.length === batchSize);

  items.forEach((item: any) => {
    engagementListMatters.push({
      ID: item.ID,
      engListID: item.ID,
      newMatterNumber: item.Title,
      creationMatterNumber: "",
      newMatterId: "",
      newMatterSiteUrl: "",
      creationMatterSiteUrl: "",
      templateType: "",
      newMatterPortalId: "",
      newMatterEngagementName: item.EngagementName,
      creation: false,
      clientNumber: item.ClientNumber,
      team: item.Team,
      newMatterWorkYear: item.WorkYear,
      creationMatterWorkYear: "",
      portalType: "",
      engagementNumberEndZero: item.WorkYear === "" ? item.Title : undefined,
      industryType: "",
      supplemental: "",
      newMatterPortalExpirationDate: createDate18MonthsFromNow().toISOString(),
      newMatterFileExpirationDate: createFileExpirationDate().toISOString(),
      isNotificationEmail: false,
      siteOwner: "",
      Portals_x0020_Created: item.Portals_x0020_Created,
    });
  });

  return { engagementListMatters };
};
