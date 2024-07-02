import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";

const hubSite = Web(GlobalValues.HubSiteURL);

interface MatterAndRolloverData {
  ID: string;
  newMatterNumber: string;
  rolloverMatterNumber: string;
  newMatterId: string;
  // rolloverMatterId: string;
  newMatterSiteUrl: string;
  rolloverMatterSiteUrl: string;
  templateType: string;
  // rolloverMatterTemplateType: string;
  newMatterPortalId: string;
  // rolloverMatterPortalId: string;
  newMatterEngagementName: string;
  // rolloverMatterEngagementName: string;
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
  isNotificationEmail: boolean;
  siteOwner: string; // needs to be email address
}



// function to create the initial date for the DatePicker component
  const createDate18MonthsFromNow = (): Date => {
    console.log("createDate18MonthsFromNow fired::");
    const date = new Date();
    date.setMonth(date.getMonth() + 18);
    return date;
  };



// TODO: Implement a function to call Engagement Portal List and get the corresponding item for the previous year for the data that is being rolled over
const getRolloverEngagementPortalItems = async (mattersData, clientNumber: string) => {

  console.log("getRolloverEngagementPortalItems firing::");
  let audRolloverArr = [];
  let taxRolloverArr = [];
  // define new object to hold the rollover data and matter number matching data
  let rolloverData: MatterAndRolloverData = {
    ID: "",
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
    isNotificationEmail: true,
    siteOwner: "",
  };

  for (const matter of mattersData) {
    let matterNumber = matter.Title;
    let workYear = matter.WorkYear;
    let previousYear: string;
    let team = matter.Team === "TAX" ? "Tax" : "Assurance";

    // subtract 1 year from the workYear value passed in
    // WORKING

    // factor logic for regular matter numbers that contain a work year
    if (matterNumber.WorkYear !== '') {
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

          console.log("Previous Portal ID::", rowData[0].PortalId);
          console.log("New Portal ID::", newPortalId);

          // factor new newMatterSiteUrl
          let siteUrl = rowData[0]['SiteUrl.desc'];

          const siteUrlSegments = siteUrl.split("/");
          siteUrlSegments[siteUrlSegments.length - 1] = newPortalId;
          siteUrl = siteUrlSegments.join("/");
          let newMatterSiteUrl = siteUrl;

          console.log("Previous Site URL::", rowData[0]["SiteUrl.desc"]);
          console.log("New Site URL::", newMatterSiteUrl);

          // assign all necessary data to the rolloverData object
          rolloverData.ID = rowData[0].ID;
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
          rolloverData.engagementNumberEndZero = rowData[0].EngagementNumberEndZero;
          rolloverData.templateType = rowData[0].TemplateType;
          rolloverData.industryType = rowData[0].IndustryType;
          rolloverData.supplemental = rowData[0].Supplemental;
          rolloverData.newMatterPortalExpirationDate = createDate18MonthsFromNow().toString(); // TODO: take user's input for newMatterPortalExpirationDate from the UI

          // console.log("Rollover Data::", rolloverData);
          taxRolloverArr.push({...rolloverData});

        } else if (rowData.length > 0 && rowData[0].Team === "Assurance") {
          audRolloverArr.push(rowData[0]);
        }



      } else {
        // factor logic for matter numbers that do not contain a work year
        let newMatterNumber = `${matterNumber}-${previousYear}`;
        console.log("New Matter Number::", newMatterNumber);
      }


    // console.table(engagementPortalList.Row);
  };


  console.log("Tax Rollover Array::", taxRolloverArr);
  console.log("Audit Rollover Array::", audRolloverArr);

  return { taxRolloverArr, audRolloverArr };

};




// Get the list of matters for the client site, filter for client number and portals created
export const getMatterNumbersForClientSite = async (
  clientSiteNumber: string
) => {
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
                        <IsNull>
                        <FieldRef Name='Portals_x0020_Created'/>
                        </IsNull>
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

  const rolloverPortalItems = await getRolloverEngagementPortalItems(engagementListMatters.Row, clientSiteNumber);

  taxMatters = rolloverPortalItems.taxRolloverArr;
  audMatters = rolloverPortalItems.audRolloverArr;

  console.log("TAX Data from getMatterNumbersForClientSite func::", taxMatters);
  console.log("AUD Data from getMatterNumbersForClientSite func::", audMatters);


  return {taxMatters, audMatters};


};
