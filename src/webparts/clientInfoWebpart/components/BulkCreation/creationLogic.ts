import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

const hubSite = Web(GlobalValues.HubSiteURL);

export interface MatterAndCreationData {
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

const getCreationEngagementPortalItems = async (
  mattersData,
  clientNumber: string
): Promise<{
  taxCreationArr: MatterAndCreationData[];
  audCreationArr: MatterAndCreationData[];
}> => {
  console.log("getCreationEngagementPortalItems firing::");
  let audCreationArr = [];
  let taxCreationArr = [];
  let creationData: MatterAndCreationData = {
    ID: "",
    engListID: "",
    newMatterNumber: "",
    creationMatterNumber: "",
    newMatterId: "",
    newMatterSiteUrl: "",
    creationMatterSiteUrl: "",
    templateType: "",
    newMatterPortalId: "",
    newMatterEngagementName: "",
    creation: true,
    clientNumber: "",
    team: "",
    newMatterWorkYear: "",
    creationMatterWorkYear: "",
    portalType: "",
    engagementNumberEndZero: "",
    industryType: "",
    supplemental: "",
    newMatterPortalExpirationDate: "",
    newMatterFileExpirationDate: createFileExpirationDate().toString(),
    isNotificationEmail: true,
    siteOwner: "",
  };

  for (const matter of mattersData) {
    let matterNumber: string = matter.Title;
    let workYear = matter.WorkYear;
    let previousYear: string;
    let team: string = matter.Team === "TAX" ? "Tax" : "Assurance";
    let engagementNumberEndZero: string;
    let portalsCreated = matter.Portals_x0020_Created;

    if (workYear === "") {
      engagementNumberEndZero = matterNumber;
    }

    if (workYear !== "" && !portalsCreated.includes('WF')) {
      previousYear = (parseInt(workYear, 10) - 1).toString();
      let lastTwoDigitsOfYear = previousYear.slice(-2);
      let matterNumberArray = matterNumber.split("-");
      let matterNumberPrefix = matterNumberArray[0];
      let matterNumberSuffix = matterNumberArray[1];
      let previousPortalYearMatterNumber = `${matterNumberPrefix}-${matterNumberSuffix}-${lastTwoDigitsOfYear}`;

      const engagementPortalList = await hubSite.lists
        .getByTitle("Engagement Portal List")
        .renderListDataAsStream({
          ViewXml: `<View>
                <Query>
                  <Where>
                    <And>
                      <And>
                        <And>
                          <And>
                          <And>
                            <Eq>
                            <FieldRef Name='ClientNumber'/>
                            <Value Type='Text'>${clientNumber}</Value>
                            </Eq>
                            <Eq>
                            <FieldRef Name='Team'/>
                            <Value Type='Text'>${team}</Value>
                            </Eq>
                          </And>
                            <Eq>
                            <FieldRef Name='WorkYear'/>
                            <Value Type='Text'>${previousYear}</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='PortalType'/>
                            <Value Type='Text'>Workflow</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='IsActive'/>
                            <Value Type='Integer'>1</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='Title'/>
                            <Value Type='Text'>${previousPortalYearMatterNumber}</Value>
                          </Eq>
                          </And>
                  </Where>
                  </Query>
                  <ViewFields>
                    <FieldRef Name='Title'/>
                    <FieldRef Name='ClientNumber'/>
                    <FieldRef Name='EngagementName'/>
                    <FieldRef Name='ID'/>
                    <FieldRef Name='WorkYear'/>
                    <FieldRef Name='Team'/>
                    <FieldRef Name='IsActive'/>
                    <FieldRef Name='EngagementNumberEndZero'/>
                    <FieldRef Name='SiteUrl'/>
                    <FieldRef Name='TemplateType'/>
                    <FieldRef Name='IndustryType'/>
                    <FieldRef Name='Supplemental'/>
                    <FieldRef Name='PortalId'/>
                    <FieldRef Name='SiteOwner'/>
                  </ViewFields>
                </View>`,
        });

      const rowData = engagementPortalList.Row;

      if (rowData.length > 0) {
        let creationPortalId = rowData[0].PortalId;
        const lastTwoDigitsOfWorkYear = matter.WorkYear.toString().slice(-2);
        let segments = creationPortalId.split("-");
        segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
        creationPortalId = segments.join("-");
        let newPortalId = creationPortalId;

        let siteUrl = rowData[0]["SiteUrl.desc"];
        const siteUrlSegments = siteUrl.split("/");
        siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
        let newMatterSiteUrl = siteUrlSegments.join("/");

        creationData = {
          ...creationData,
          ID: rowData[0].ID,
          engListID: matter.ID,
          newMatterNumber: matterNumber,
          creationMatterNumber: previousPortalYearMatterNumber,
          newMatterId: rowData[0].ID,
          newMatterSiteUrl: newMatterSiteUrl,
          creationMatterSiteUrl: rowData[0]["SiteUrl.desc"],
          newMatterPortalId: newPortalId,
          newMatterEngagementName: matter.EngagementName,
          clientNumber: clientNumber,
          team: rowData[0].Team,
          newMatterWorkYear: workYear,
          creationMatterWorkYear: previousYear,
          portalType: rowData[0].PortalType,
          engagementNumberEndZero: rowData[0].EngagementNumberEndZero,
          templateType: rowData[0].TemplateType,
          industryType: rowData[0].IndustryType,
          supplemental: rowData[0].Supplemental,
          newMatterPortalExpirationDate: createDate18MonthsFromNow().toString(),
        };

        if (rowData[0].Team === "Tax") {
          taxCreationArr.push({ ...creationData });
        } else if (rowData[0].Team === "Assurance") {
          audCreationArr.push({ ...creationData });
        }
      }
    } else if (workYear === "" && portalsCreated.includes("WF")) {
      const engagementPortalList = await hubSite.lists
        .getByTitle("Engagement Portal List")
        .renderListDataAsStream({
          ViewXml: `<View>
                <Query>
                  <Where>
                      <And>
                        <And>
                          <And>
                          <And>
                            <Eq>
                            <FieldRef Name='ClientNumber'/>
                            <Value Type='Text'>${clientNumber}</Value>
                            </Eq>
                            <Eq>
                            <FieldRef Name='Team'/>
                            <Value Type='Text'>${team}</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='PortalType'/>
                            <Value Type='Text'>Workflow</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='IsActive'/>
                            <Value Type='Integer'>1</Value>
                            </Eq>
                          </And>
                          <Eq>
                            <FieldRef Name='EngagementNumberEndZero'/>
                            <Value Type='Text'>${engagementNumberEndZero}</Value>
                          </Eq>
                          </And>
                  </Where>
                  </Query>
                  <ViewFields>
                    <FieldRef Name='Title'/>
                    <FieldRef Name='ClientNumber'/>
                    <FieldRef Name='EngagementName'/>
                    <FieldRef Name='ID'/>
                    <FieldRef Name='WorkYear'/>
                    <FieldRef Name='Team'/>
                    <FieldRef Name='IsActive'/>
                    <FieldRef Name='EngagementNumberEndZero'/>
                    <FieldRef Name='SiteUrl'/>
                    <FieldRef Name='TemplateType'/>
                    <FieldRef Name='IndustryType'/>
                    <FieldRef Name='Supplemental'/>
                    <FieldRef Name='PortalId'/>
                  </ViewFields>
                </View>`,
        });

      const rowData = engagementPortalList.Row;
      const latestYear = rowData.reduce(
        (max, obj) => (obj.WorkYear > max ? obj.WorkYear : max),
        rowData[0].WorkYear
      );

      const latestYearPortalItem = rowData.filter(
        (obj) => obj.WorkYear === latestYear
      );

      let creationPortalId = rowData[0].PortalId;
      let incrementedWorkYear = parseInt(latestYearPortalItem[0].WorkYear) + 1;
      const lastTwoDigitsOfWorkYear = incrementedWorkYear.toString().slice(-2);
      let segments = creationPortalId.split("-");
      segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
      let newPortalId = segments.join("-");
      let matterNumberSegments = engagementNumberEndZero.split("-");
      matterNumberSegments[matterNumberSegments.length - 1] = lastTwoDigitsOfWorkYear;
      let updatedMatterNumber = matterNumberSegments.join("-");
      let siteUrl = latestYearPortalItem[0]["SiteUrl.desc"];
      const siteUrlSegments = siteUrl.split("/");
      siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
      let newMatterSiteUrl = siteUrlSegments.join("/");

      creationData = {
        ...creationData,
        ID: latestYearPortalItem[0].ID,
        engListID: matter.ID,
        newMatterNumber: updatedMatterNumber,
        creationMatterNumber: latestYearPortalItem[0].Title,
        newMatterId: latestYearPortalItem[0].ID,
        newMatterSiteUrl: newMatterSiteUrl,
        creationMatterSiteUrl: latestYearPortalItem[0]["SiteUrl.desc"],
        newMatterPortalId: newPortalId,
        newMatterEngagementName: matter.EngagementName,
        clientNumber: clientNumber,
        team: latestYearPortalItem[0].Team,
        newMatterWorkYear: incrementedWorkYear.toString(),
        creationMatterWorkYear: latestYearPortalItem[0].WorkYear,
        portalType: latestYearPortalItem[0].PortalType,
        engagementNumberEndZero: latestYearPortalItem[0].EngagementNumberEndZero,
        templateType: latestYearPortalItem[0].TemplateType,
        industryType: latestYearPortalItem[0].IndustryType,
        supplemental: latestYearPortalItem[0].Supplemental,
        newMatterPortalExpirationDate: createDate18MonthsFromNow().toString(),
      };

      if (latestYearPortalItem[0].Team === "Tax") {
        taxCreationArr.push({ ...creationData });
      } else if (latestYearPortalItem[0].Team === "Assurance") {
        audCreationArr.push({ ...creationData });
      }
    }
  }

  return { taxCreationArr, audCreationArr };
};

export const getMatterNumbersForClientSite = async (
  clientSiteNumber: string
): Promise<{
  taxMatters: MatterAndCreationData[];
  audMatters: MatterAndCreationData[];
}> => {
  let taxMatters = [];
  let audMatters = [];

  console.log("getMatterNumbersForClientSite firing::", clientSiteNumber);

  let engagementListMatters = await hubSite.lists
    .getByTitle("Engagement List")
    .renderListDataAsStream({
      ViewXml: `<View>
                <Query>
                  <Where>
                      <And>
                      <Eq>
                      <FieldRef Name='ClientNumber'/>
                      <Value Type='Text'>${clientSiteNumber}</Value>
                      </Eq>
                      <Neq>
                        <FieldRef Name='Team'/>
                        <Value Type='Text'>ADV</Value>
                      </Neq>
                      </And>
                  </Where>
                </Query>
                <ViewFields>
                  <FieldRef Name='Title'/>
                  <FieldRef Name='ClientNumber'/>
                  <FieldRef Name='EngagementName'/>
                  <FieldRef Name='ID'/>
                  <FieldRef Name='WorkYear'/>
                  <FieldRef Name='Team'/>
                  <FieldRef Name='Portals_x0020_Created'/>
                </ViewFields>
              </View>`,
    });

  console.table(engagementListMatters.Row);

  const creationPortalItems = await getCreationEngagementPortalItems(
    engagementListMatters.Row,
    clientSiteNumber
  );

  taxMatters = creationPortalItems.taxCreationArr;
  audMatters = creationPortalItems.audCreationArr;

  return { taxMatters, audMatters };
};
