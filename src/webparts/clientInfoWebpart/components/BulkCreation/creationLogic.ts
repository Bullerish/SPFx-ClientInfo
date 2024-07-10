import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

const hubSite = Web(GlobalValues.HubSiteURL);

export interface MatterAndCreationData {
  ID: string;
  newMatterNumber: string;
  creationMatterNumber: string;
  newMatterId: string;
  // creationMatterId: string;
  newMatterSiteUrl: string;
  creationMatterSiteUrl: string;
  templateType: string;
  // creationMatterTemplateType: string;
  newMatterPortalId: string;
  // creationMatterPortalId: string;
  newMatterEngagementName: string;
  // creationMatterEngagementName: string;
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
  isNotificationEmail: boolean;
  siteOwner: string | ISiteUserInfo | number; // needs to be email address
}

// function to create the initial date for the DatePicker component
export const createDate18MonthsFromNow = (): Date => {
  // console.log("createDate18MonthsFromNow fired::");
  const date = new Date();
  date.setMonth(date.getMonth() + 18);
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
  // define new object to hold the creation data and matter number matching data
  let creationData: MatterAndCreationData = {
    ID: "",
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

    // subtract 1 year from the workYear value passed in
    // WORKING

    // factor logic for regular matter numbers that contain a work year
    if (workYear !== "" && portalsCreated === "") {
      previousYear = (parseInt(workYear, 10) - 1).toString();
      let lastTwoDigitsOfYear = previousYear.slice(-2);
      let matterNumberArray = matterNumber.split("-");
      let matterNumberPrefix = matterNumberArray[0];
      let matterNumberSuffix = matterNumberArray[1];
      let previousPortalYearMatterNumber = `${matterNumberPrefix}-${matterNumberSuffix}-${lastTwoDigitsOfYear}`;

      // call the engagement portal list and get the corresponding item for the previous year based on the incoming matter number
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

      // console.log('Matter Number Title::', matter.Title);
      // console.log('Matter Number Work Year::', workYear);
      // console.log('Team::', matter.Team);
      // console.log("Previous Portal Year Matter Number::", previousPortalYearMatterNumber);
      // console.log(engagementPortalList.Row);

      if (rowData.length > 0 && rowData[0].Team === "Tax") {
        // taxCreationArr.push(rowData[0]);
        // console.log(engagementPortalList.Row);

        // TODO: configure and build newMatterSiteUrl, newMatterPortalId, and in BulkCreation.tsx factor the siteOwner and PortalExpirationDate from the UI

        // factor newMatterSiteUrl
        let creationPortalId = rowData[0].PortalId;
        const lastTwoDigitsOfWorkYear = matter.WorkYear.toString().slice(-2);
        let segments = creationPortalId.split("-");

        segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
        creationPortalId = segments.join("-");
        let newPortalId = creationPortalId;

        // console.log("Previous Portal ID::", rowData[0].PortalId);
        // console.log("New Portal ID::", newPortalId);

        // factor new newMatterSiteUrl
        let siteUrl = rowData[0]["SiteUrl.desc"];

        const siteUrlSegments = siteUrl.split("/");
        siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
        siteUrl = siteUrlSegments.join("/");
        let newMatterSiteUrl = siteUrl;

        // console.log("Previous Site URL::", rowData[0]["SiteUrl.desc"]);
        // console.log("New Site URL::", newMatterSiteUrl);

        // assign all necessary data to the creationData object
        creationData.ID = rowData[0].ID;
        creationData.newMatterNumber = matterNumber;
        creationData.creationMatterNumber = previousPortalYearMatterNumber;
        creationData.newMatterId = rowData[0].ID;
        creationData.newMatterSiteUrl = newMatterSiteUrl;
        creationData.creationMatterSiteUrl = rowData[0]["SiteUrl.desc"];
        creationData.newMatterPortalId = newPortalId;
        creationData.newMatterEngagementName = matter.EngagementName;
        creationData.clientNumber = clientNumber;
        creationData.team = rowData[0].Team;
        creationData.newMatterWorkYear = workYear;
        creationData.creationMatterWorkYear = previousYear;
        creationData.portalType = rowData[0].PortalType;
        creationData.engagementNumberEndZero =
          rowData[0].EngagementNumberEndZero;
        creationData.templateType = rowData[0].TemplateType;
        creationData.industryType = rowData[0].IndustryType;
        creationData.supplemental = rowData[0].Supplemental;
        creationData.newMatterPortalExpirationDate =
          createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI

        // console.log("Creation Data::", creationData);
        taxCreationArr.push({ ...creationData });

        // check if the rowData is not empty and the team is Assurance
      } else if (rowData.length > 0 && rowData[0].Team === "Assurance") {
        // factor newMatterSiteUrl
        let creationPortalId = rowData[0].PortalId;
        const lastTwoDigitsOfWorkYear = matter.WorkYear.toString().slice(-2);
        let segments = creationPortalId.split("-");
        segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
        creationPortalId = segments.join("-");
        let newPortalId = creationPortalId;

        // console.log("Previous Portal ID::", rowData[0].PortalId);
        // console.log("New Portal ID::", newPortalId);

        // factor new newMatterSiteUrl
        let siteUrl = rowData[0]["SiteUrl.desc"];

        const siteUrlSegments = siteUrl.split("/");
        siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
        siteUrl = siteUrlSegments.join("/");
        let newMatterSiteUrl = siteUrl;

        // console.log("Previous Site URL::", rowData[0]["SiteUrl.desc"]);
        // console.log("New Site URL::", newMatterSiteUrl);

        // assign all necessary data to the creationData object
        creationData.ID = rowData[0].ID;
        creationData.newMatterNumber = matterNumber;
        creationData.creationMatterNumber = previousPortalYearMatterNumber;
        creationData.newMatterId = rowData[0].ID;
        creationData.newMatterSiteUrl = newMatterSiteUrl;
        creationData.creationMatterSiteUrl = rowData[0]["SiteUrl.desc"];
        creationData.newMatterPortalId = newPortalId;
        creationData.newMatterEngagementName = matter.EngagementName;
        creationData.clientNumber = clientNumber;
        creationData.team = rowData[0].Team;
        creationData.newMatterWorkYear = workYear;
        creationData.creationMatterWorkYear = previousYear;
        creationData.portalType = rowData[0].PortalType;
        creationData.engagementNumberEndZero =
          rowData[0].EngagementNumberEndZero;
        creationData.templateType = rowData[0].TemplateType;
        creationData.industryType = rowData[0].IndustryType;
        creationData.supplemental = rowData[0].Supplemental;
        creationData.newMatterPortalExpirationDate =
          createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI

        audCreationArr.push({ ...creationData });
      }
    } else if (workYear === "" && portalsCreated.includes("WF") === true) {
      // console.log("Inside -00 else block::");
      // console.log("logging engagementNumberEndZero::", engagementNumberEndZero);

      // call the engagement portal list and get the corresponding item for the previous year based on the incoming matter number
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
      // console.log("logging rowData for -00 matters::", rowData);

      const latestYear = rowData.reduce(
        (max, obj) => (obj.WorkYear > max ? obj.WorkYear : max),
        rowData[0].WorkYear
      );

      const latestYearPortalItem = rowData.filter(
        (obj) => obj.WorkYear === latestYear
      );

      // console.log("Most recent portal by year::", latestYearPortalItem);

      // factor logic for matter numbers that do not contain a work year (-00)
      // factor newMatterSiteUrl
      let creationPortalId = rowData[0].PortalId;

      let incrementedWorkYear = parseInt(latestYearPortalItem[0].WorkYear) + 1;

      // console.log('logging incrementedWorkYear::', incrementedWorkYear);

      const lastTwoDigitsOfWorkYear =
      incrementedWorkYear.toString().slice(-2);
      let segments = creationPortalId.split("-");
      segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
      creationPortalId = segments.join("-");

      let newPortalId = creationPortalId;

      // TODO: create new matter number based on the engagement number end zero and the incremented work year
      let matterNumberSegments = engagementNumberEndZero.split("-");
      matterNumberSegments[matterNumberSegments.length - 1] = lastTwoDigitsOfWorkYear;
      let updatedMatterNumber = matterNumberSegments.join("-");

      // console.log("Previous Portal ID::", rowData[0].PortalId);
      // console.log("New Portal ID::", newPortalId);

      // factor new newMatterSiteUrl
      let siteUrl = latestYearPortalItem[0]["SiteUrl.desc"];

      const siteUrlSegments = siteUrl.split("/");
      siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
      siteUrl = siteUrlSegments.join("/");
      let newMatterSiteUrl = siteUrl;

      // console.log("Previous Site URL::", rowData[0]["SiteUrl.desc"]);
      // console.log("New Site URL::", newMatterSiteUrl);

      // assign all necessary data to the creationData object
      creationData.ID = latestYearPortalItem[0].ID;
      creationData.newMatterNumber = updatedMatterNumber;
      creationData.creationMatterNumber = latestYearPortalItem[0].Title;
      creationData.newMatterId = latestYearPortalItem[0].ID;
      creationData.newMatterSiteUrl = newMatterSiteUrl;
      creationData.creationMatterSiteUrl =
        latestYearPortalItem[0]["SiteUrl.desc"];
      creationData.newMatterPortalId = newPortalId;
      creationData.newMatterEngagementName = matter.EngagementName;
      creationData.clientNumber = clientNumber;
      creationData.team = latestYearPortalItem[0].Team;
      creationData.newMatterWorkYear = incrementedWorkYear.toString();
      creationData.creationMatterWorkYear = latestYearPortalItem[0].WorkYear;
      creationData.portalType = latestYearPortalItem[0].PortalType;
      creationData.engagementNumberEndZero = rowData[0].EngagementNumberEndZero;
      creationData.templateType = latestYearPortalItem[0].TemplateType;
      creationData.industryType = latestYearPortalItem[0].IndustryType;
      creationData.supplemental = latestYearPortalItem[0].Supplemental;
      creationData.newMatterPortalExpirationDate =
        createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI


      // console.log("-00 Creation Data::", {...creationData});

      if (latestYearPortalItem[0].Team === "Tax") {
        taxCreationArr.push({ ...creationData });
      } else if (latestYearPortalItem[0].Team === "Assurance") {
        audCreationArr.push({ ...creationData });
      }


      // audCreationArr.push({ ...creationData });
    }


  }

  // console.log("Tax Creation Array::", taxCreationArr);
  // console.log("Audit Creation Array::", audCreationArr);

  return { taxCreationArr, audCreationArr };
};


// Get the list of matters for the client site, filter for client number and portals created
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

  // console.table(engagementListMatters.Row);

  const creationPortalItems = await getCreationEngagementPortalItems(
    engagementListMatters.Row,
    clientSiteNumber
  );

  taxMatters = creationPortalItems.taxCreationArr;
  audMatters = creationPortalItems.audCreationArr;

  // console.log("TAX Data from getMatterNumbersForClientSite func::", taxMatters);
  // console.log("AUD Data from getMatterNumbersForClientSite func::", audMatters);

  return { taxMatters, audMatters };
};
