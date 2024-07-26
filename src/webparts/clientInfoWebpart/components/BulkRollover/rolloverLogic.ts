import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

const hubSite = Web(GlobalValues.HubSiteURL);

export interface MatterAndRolloverData {
  ID: string;
  engListID: string;
  newMatterNumber: string;
  rolloverMatterNumber: string;
  newMatterId: string;
  // rolloverMatterId: string;
  newMatterSiteUrl: string;
  rolloverMatterSiteUrl: string;
  templateType: string;
  newMatterPortalId: string;
  newMatterEngagementName: string;
  rollover: boolean;
  clientNumber: string;
  team: string;
  newMatterWorkYear: string;
  rolloverMatterWorkYear: string;
  portalType: string;
  engagementNumberEndZero?: string;
  industryType?: string;
  supplemental: string;
  newMatterPortalExpirationDate: string;
  newMatterFileExpirationDate: string;
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

// func to create the initial date for the file expiration date (12 months from now)
const createFileExpirationDate = (): Date => {
  // console.log("createDate18MonthsFromNow fired::");
  const date = new Date();
  date.setMonth(date.getMonth() + 12);
  return date;
};

// bulk of the logic to get the rollover data for the client site. Differences for -00 matters and regular matters
const getRolloverEngagementPortalItems = async (
  mattersData,
  clientNumber: string
): Promise<{
  taxRolloverArr: MatterAndRolloverData[];
  audRolloverArr: MatterAndRolloverData[];
}> => {
  console.log("getRolloverEngagementPortalItems firing::");
  let audRolloverArr = [];
  let taxRolloverArr = [];
  // define new object to hold the rollover data and matter number matching data
  let rolloverData: MatterAndRolloverData = {
    ID: "",
    engListID: "",
    newMatterNumber: "",
    rolloverMatterNumber: "",
    newMatterId: "",
    newMatterSiteUrl: "",
    rolloverMatterSiteUrl: "",
    templateType: "",
    newMatterPortalId: "",
    newMatterEngagementName: "",
    rollover: true,
    clientNumber: "",
    team: "",
    newMatterWorkYear: "",
    rolloverMatterWorkYear: "",
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

    // subtract 1 year from the workYear value passed in
    // WORKING

    // factor logic for regular matter numbers that contain a work year
    if (workYear !== "" && !portalsCreated.includes('WF')) {
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
        // console.log("logging rowData for regular matters::", rowData);

      if (rowData.length > 0 && rowData[0].Team === "Tax") {
        // taxRolloverArr.push(rowData[0]);
        // console.log(engagementPortalList.Row);

        // TODO: configure and build newMatterSiteUrl, newMatterPortalId, and in BulkRollover.tsx factor the siteOwner and PortalExpirationDate from the UI

        // factor newMatterSiteUrl
        let rolloverPortalId = rowData[0].PortalId;
        const lastTwoDigitsOfWorkYear = matter.WorkYear.toString().slice(-2);
        let segments = rolloverPortalId.split("-");

        segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
        rolloverPortalId = segments.join("-");
        let newPortalId = rolloverPortalId;

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

        // assign all necessary data to the rolloverData object
        rolloverData.ID = rowData[0].ID;
        rolloverData.engListID = matter.ID;
        rolloverData.newMatterNumber = matterNumber;
        rolloverData.rolloverMatterNumber = previousPortalYearMatterNumber;
        rolloverData.newMatterId = rowData[0].ID;
        rolloverData.newMatterSiteUrl = newMatterSiteUrl;
        rolloverData.rolloverMatterSiteUrl = rowData[0]["SiteUrl.desc"];
        rolloverData.newMatterPortalId = newPortalId;
        rolloverData.newMatterEngagementName = matter.EngagementName;
        rolloverData.clientNumber = clientNumber;
        rolloverData.team = rowData[0].Team;
        rolloverData.newMatterWorkYear = workYear;
        rolloverData.rolloverMatterWorkYear = previousYear;
        rolloverData.portalType = rowData[0].PortalType;
        rolloverData.engagementNumberEndZero =
          rowData[0].EngagementNumberEndZero;
        rolloverData.templateType = rowData[0].TemplateType;
        rolloverData.industryType = rowData[0].IndustryType;
        rolloverData.supplemental = rowData[0].Supplemental;
        rolloverData.newMatterPortalExpirationDate =
          createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI

        // console.log("Rollover Data::", rolloverData);
        taxRolloverArr.push({ ...rolloverData });

        // check if the rowData is not empty and the team is Assurance
      } else if (rowData.length > 0 && rowData[0].Team === "Assurance") {
        // factor newMatterSiteUrl
        let rolloverPortalId = rowData[0].PortalId;
        const lastTwoDigitsOfWorkYear = matter.WorkYear.toString().slice(-2);
        let segments = rolloverPortalId.split("-");
        segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
        rolloverPortalId = segments.join("-");
        let newPortalId = rolloverPortalId;

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

        // assign all necessary data to the rolloverData object
        rolloverData.ID = rowData[0].ID;
        rolloverData.engListID = matter.ID;
        rolloverData.newMatterNumber = matterNumber;
        rolloverData.rolloverMatterNumber = previousPortalYearMatterNumber;
        rolloverData.newMatterId = rowData[0].ID;
        rolloverData.newMatterSiteUrl = newMatterSiteUrl;
        rolloverData.rolloverMatterSiteUrl = rowData[0]["SiteUrl.desc"];
        rolloverData.newMatterPortalId = newPortalId;
        rolloverData.newMatterEngagementName = matter.EngagementName;
        rolloverData.clientNumber = clientNumber;
        rolloverData.team = rowData[0].Team;
        rolloverData.newMatterWorkYear = workYear;
        rolloverData.rolloverMatterWorkYear = previousYear;
        rolloverData.portalType = rowData[0].PortalType;
        rolloverData.engagementNumberEndZero =
          rowData[0].EngagementNumberEndZero;
        rolloverData.templateType = rowData[0].TemplateType;
        rolloverData.industryType = rowData[0].IndustryType;
        rolloverData.supplemental = rowData[0].Supplemental;
        rolloverData.newMatterPortalExpirationDate =
          createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI

        audRolloverArr.push({ ...rolloverData });



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
      // console.log("logging rowData for -00 matters::", rowData.length);

      if (rowData.length > 0) {

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
      let rolloverPortalId = rowData[0].PortalId;

      let incrementedWorkYear = parseInt(latestYearPortalItem[0].WorkYear) + 1;

      // console.log('logging incrementedWorkYear::', incrementedWorkYear);

      const lastTwoDigitsOfWorkYear =
      incrementedWorkYear.toString().slice(-2);
      let segments = rolloverPortalId.split("-");
      segments[segments.length - 1] = lastTwoDigitsOfWorkYear;
      rolloverPortalId = segments.join("-");

      let newPortalId = rolloverPortalId;

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

      // assign all necessary data to the rolloverData object
      rolloverData.ID = latestYearPortalItem[0].ID;
      rolloverData.engListID = matter.ID;
      rolloverData.newMatterNumber = updatedMatterNumber;
      rolloverData.rolloverMatterNumber = latestYearPortalItem[0].Title;
      rolloverData.newMatterId = latestYearPortalItem[0].ID;
      rolloverData.newMatterSiteUrl = newMatterSiteUrl;
      rolloverData.rolloverMatterSiteUrl =
        latestYearPortalItem[0]["SiteUrl.desc"];
      rolloverData.newMatterPortalId = newPortalId;
      rolloverData.newMatterEngagementName = matter.EngagementName;
      rolloverData.clientNumber = clientNumber;
      rolloverData.team = latestYearPortalItem[0].Team;
      rolloverData.newMatterWorkYear = incrementedWorkYear.toString();
      rolloverData.rolloverMatterWorkYear = latestYearPortalItem[0].WorkYear;
      rolloverData.portalType = latestYearPortalItem[0].PortalType;
      rolloverData.engagementNumberEndZero = rowData[0].EngagementNumberEndZero;
      rolloverData.templateType = latestYearPortalItem[0].TemplateType;
      rolloverData.industryType = latestYearPortalItem[0].IndustryType;
      rolloverData.supplemental = latestYearPortalItem[0].Supplemental;
      rolloverData.newMatterPortalExpirationDate =
        createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI


      // console.log("-00 Rollover Data::", {...rolloverData});

      if (latestYearPortalItem[0].Team === "Tax") {
        taxRolloverArr.push({ ...rolloverData });
      } else if (latestYearPortalItem[0].Team === "Assurance") {
        audRolloverArr.push({ ...rolloverData });
      }

    }
      // audRolloverArr.push({ ...rolloverData });
    }


  }

  // console.log("Tax Rollover Array::", taxRolloverArr);
  // console.log("Audit Rollover Array::", audRolloverArr);

  return { taxRolloverArr, audRolloverArr };
};

// Get the list of matters for the client site, filter for client number and portals created
export const getMatterNumbersForClientSite = async (
  clientSiteNumber: string
): Promise<{
  taxMatters: MatterAndRolloverData[];
  audMatters: MatterAndRolloverData[];
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

  const rolloverPortalItems = await getRolloverEngagementPortalItems(
    engagementListMatters.Row,
    clientSiteNumber
  );

  taxMatters = rolloverPortalItems.taxRolloverArr;
  audMatters = rolloverPortalItems.audRolloverArr;

  // console.log("TAX Data from getMatterNumbersForClientSite func::", taxMatters);
  // console.log("AUD Data from getMatterNumbersForClientSite func::", audMatters);

  return { taxMatters, audMatters };
};
