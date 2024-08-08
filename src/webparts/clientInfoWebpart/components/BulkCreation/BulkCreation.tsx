import * as React from 'react';
import { useState, useEffect, useLayoutEffect } from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import {
  DefaultButton, PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import {
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react/lib/MessageBar";
import { Web } from '@pnp/sp/webs';
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import {
  DatePicker,
} from "office-ui-fabric-react/lib/DatePicker";
import {
  PeoplePicker, PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  ListView, IViewField, SelectionMode,
} from "@pnp/spfx-controls-react/lib/ListView";
import { getMatterNumbersForClientSite, MatterAndCreationData, createDate18MonthsFromNow } from './creationLogic';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { ClientInfoClass } from '../../Dataprovider/ClientInfoClass'; // Make sure this import path is correct
import { set } from '@microsoft/sp-lodash-subset';

const BulkCreation = ({
  spContext,
  isBulkCreationOpen,
  onBulkCreationModalHide,
}): React.ReactElement => {
  const [team, setTeam] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [teamKey, setTeamKey] = useState<string>("");
  const [selectedDates, setSelectedDates] = useState({});
  const [portalType, setPortalType] = useState<string>("");
  const [isDataLoaded, setIsDataLoaded] = useState<boolean>(false);
  const [items, setItems] = useState<MatterAndCreationData[]>([]);
  const [itemsStaged, setItemsStaged] = useState<MatterAndCreationData[]>([]);
  const [portalSelected, setPortalSelected] = useState([]);
  const [enableNextButton, setEnableNextButton] = useState<boolean>(false);
  const [isConfirmationScreen, setIsConfirmationScreen] = useState<boolean>(false);
  const [isDataSubmitted, setIsDataSubmitted] = useState<boolean>(false);
  const [industryTypes, setIndustryTypes] = useState<any[]>([]);
  const [supplementals, setSupplementals] = useState<any[]>([]);
  const [templateTypes, setTemplateTypes] = useState<any[]>([]);
  const [isTeamAndPortalDisabled, setIsTeamAndPortalDisabled] = useState<boolean>(false);
  const clientSiteAbsoluteUrl = spContext._pageContext._web.absoluteUrl;
  const clientSiteServerRelativeUrl = spContext._pageContext._web.serverRelativeUrl;
  const relativeUrlArr = clientSiteServerRelativeUrl.split("/");
  const clientSiteNumber = relativeUrlArr[relativeUrlArr.length - 1];
  const hubSite = Web(GlobalValues.HubSiteURL);

  const resetState = () => {
    onBulkCreationModalHide(false);
    setTeam("");
    setError("");
    setTeamKey("");
    setPortalType("");
    setIsDataLoaded(false);
    setItems([]);
    setItemsStaged([]);
    setPortalSelected([]);
    setEnableNextButton(false);
    setIsConfirmationScreen(false);
    setIsDataSubmitted(false);
    setIndustryTypes([]);
    setSupplementals([]);
    setTemplateTypes([]);
    setIsTeamAndPortalDisabled(false);
  };

  const viewFields: IViewField[] = [
    {
      name: "newMatterEngagementName",
      displayName: "Engagement Name",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
    },
    {
      name: "newMatterNumber",
      displayName: "Matter #",
      sorting: false,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
    },
  ];
  const confirmationViewFields: IViewField[] = [
    {
      name: "newMatterEngagementName",
      displayName: "Engagement Name",
      sorting: false,
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      name: "newMatterNumber",
      displayName: "Matter #",
      sorting: false,
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      name: "newMatterWorkYear",
      displayName: "Year",
      sorting: false,
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    ...(team === "AUD" && portalType === "workflow" ? [
      {
        name: "industryType",
        displayName: "Industry Type",
        sorting: false,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        render: (rowItem, index, column) => (
          <span>{rowItem.industryType}</span>
        ),
      },
      {
        name: "supplemental",
        displayName: "Supplemental",
        sorting: false,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        render: (rowItem, index, column) => (
          <span>{rowItem.supplemental}</span>
        ),
      },
    ] : []),
    ...(team === "TAX" && portalType === "workflow" ? [
      {
        name: "templateType",
        displayName: "Template Type",
        sorting: false,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        render: (rowItem, index, column) => (
          <span>{rowItem.templateType}</span>
        ),
      },
      {
        name: "industryType",
        displayName: "Industry Type",
        sorting: false,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        render: (rowItem, index, column) => (
          <span>{rowItem.industryType}</span>
        ),
      },
    ] : []),
    {
      name: "siteOwner",
      displayName: "Site Owner",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
      render: (rowItem, index, column) => (
        <span>{rowItem["siteOwner.Title"]}</span>
      ),
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Portal Expiration Date",
      sorting: false,
      minWidth: 150,
      maxWidth: 300,
      isResizable: false,
      render: (rowItem, index, column) => {
        return (
          <span>{new Date(rowItem.newMatterPortalExpirationDate).toLocaleDateString()}</span>
        );
      },
    },
  ];



  const onTeamChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    const selectedTeam = option.key.toLowerCase();
    let teamCode = "";
    if (selectedTeam === "assurance") {
      teamCode = "AUD";
    } else if (selectedTeam === "advisory") {
      teamCode = "ADV";
    } else if (selectedTeam === "tax") {
      teamCode = "TAX";
    }
    setTeam(teamCode);
    setTeamKey(selectedTeam);
  };

  const onPortalTypeChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    setPortalType(option.key);
  };

  const onFormatDate = (date: Date): string => {
    return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  };

  const onSelectDate = (date: Date | null | undefined, rowItemToUpdate: MatterAndCreationData): void => {
    const updatedItemsStaged = itemsStaged.map((item) => {
      if (item.ID === rowItemToUpdate.ID) {
        return {
          ...item,
          newMatterPortalExpirationDate: date ? date.toISOString() : null,
        };
      }
      return item;
    });
    setItemsStaged(updatedItemsStaged);
  };

  const getPeoplePickerItems = (itemsArr: any[], itemRow: MatterAndCreationData) => {
    const currSite = Web(GlobalValues.HubSiteURL);
    itemsArr.forEach((e) => {
      currSite.siteUsers.getByLoginName(e.loginName).get().then((user) => {
        const updatedItemsStaged = itemsStaged.map((item) => {
          if (item.ID === itemRow.ID) {
            return {
              ...item,
              siteOwner: user,
            };
          }
          return item;
        });
        setItemsStaged(updatedItemsStaged);
      });
    });
  };

  const validateSiteOwner = (itemsSiteOwner: any[], rowItemToUpdate) => {
    if (itemsSiteOwner.length > 0) {
      let userEmail = itemsSiteOwner[0].secondaryText.toLowerCase();
      if (userEmail.includes("cohnreznick.com") || userEmail.includes("cohnreznickdev")) {
        getPeoplePickerItems(itemsSiteOwner, rowItemToUpdate);
      }
    } else {
      const updatedItemsStaged = itemsStaged.map((item) => {
        if (item.ID === rowItemToUpdate.ID) {
          return {
            ...item,
            siteOwner: "",
          };
        }
        return item;
      });
      setItemsStaged(updatedItemsStaged);
    }
  };

  const moveSelectedToStaged = () => {
    const currentYearLastTwoDigits = new Date().getFullYear().toString().slice(-2);

    portalSelected.forEach((selectedItem) => {
      const newItems = items.filter((item) => item.ID !== selectedItem.ID);
      setItems(newItems);

      const isAlreadyStaged = itemsStaged.some((item) => item.ID === selectedItem.ID);
      if (!isAlreadyStaged) {
        let newMatterNumber = selectedItem.newMatterNumber;
        let engagementNumberEndZero = selectedItem.engagementNumberEndZero;
        let currentYear = new Date().getFullYear().toString();
        if (selectedItem.newMatterNumber.endsWith("00")) {
          newMatterNumber = selectedItem.newMatterNumber.slice(0, -2) + currentYearLastTwoDigits;
          engagementNumberEndZero = selectedItem.newMatterNumber;
          selectedItem.newMatterWorkYear = currentYear;
        }

        const updatedSelectedItem = {
          ...selectedItem,
          newMatterNumber,
          engagementNumberEndZero,
        };

        setItemsStaged((prevItemsStaged) => [...prevItemsStaged, updatedSelectedItem]);
      }
    });
  };

  const unstageItem = (ev, itemRowToRemove) => {
    const newItemsStaged = itemsStaged.filter((item) => item.ID !== itemRowToRemove.ID);
    setItemsStaged(newItemsStaged);
    const isAlreadyInItems = items.some((item) => item.ID === itemRowToRemove.ID);
    if (!isAlreadyInItems) {
      itemRowToRemove.siteOwner = "";
      if (itemRowToRemove.engagementNumberEndZero) {
        itemRowToRemove.newMatterNumber = itemRowToRemove.engagementNumberEndZero;
      }
      setItems((prevItems) => [...prevItems, itemRowToRemove]);
    }
  };

  const checkItemsStagedForSiteOwner = () => {
    return itemsStaged.every((item) => item.siteOwner && item.siteOwner !== "");
  };

  const submitPortalCreationData = () => {
    let mattersToUpdatePC = [];
    let selectedPortalTypeShortHand = "";
    Promise.all(itemsStaged.map(stagedItem => {
      if (stagedItem.engagementNumberEndZero === "" || stagedItem.engagementNumberEndZero === undefined) {
        mattersToUpdatePC.push(stagedItem.engListID);
      }
      let selectedTeamName;
      if (team === "AUD") {
        selectedTeamName = "Assurance";
      } else if (team === "TAX") {
        selectedTeamName = "Tax";
      } else if (team === "ADV") {
        selectedTeamName = "Advisory";
      }
      let selectedPortalType = portalType === "workflow" ? "Workflow" : "File Exchange";
      selectedPortalTypeShortHand = portalType === "workflow" ? "WF" : "FE";
      // Construct the newMatterSiteUrl and PortalId
      const newMatterSiteUrl = `${GlobalValues.SiteURL}/${team}-${selectedPortalTypeShortHand}-${stagedItem.newMatterNumber}`;
      const portalId = `${team}-${selectedPortalTypeShortHand}-${stagedItem.newMatterNumber}`;

      const itemData = {
        Title: stagedItem.newMatterNumber,
        EngagementName: stagedItem.newMatterEngagementName,
        ClientNumber: stagedItem.clientNumber,
        EngagementNumberEndZero: stagedItem.engagementNumberEndZero,
        WorkYear: stagedItem.newMatterWorkYear,
        Team: selectedTeamName,
        PortalType: selectedPortalType,
        SiteUrl: {
          __metadata: { type: "SP.FieldUrlValue" },
          Description: newMatterSiteUrl,
          Url: newMatterSiteUrl,
        },
        PortalId: portalId,
        TemplateType: stagedItem.templateType,
        IndustryType: stagedItem.industryType,
        Supplemental: stagedItem.supplemental,
        SiteOwnerId: stagedItem.siteOwner["Id"],
        PortalExpiration: new Date(stagedItem.newMatterPortalExpirationDate),
        FileExpiration: new Date(stagedItem.newMatterFileExpirationDate),
        isNotificationEmail: true,
      };
      return hubSite.lists.getByTitle("Engagement Portal List").items.add(itemData);
    }))
      .then((results) => {
        setIsDataSubmitted(true);
        if (mattersToUpdatePC.length > 0) {
          mattersToUpdatePC.forEach((matterToUpdate) => {
            updateEngListRegularMatter(matterToUpdate, selectedPortalTypeShortHand);
          });
        }
        selectedPortalTypeShortHand = "";
      })
      .catch((err) => { // Renamed 'error' to 'err'
        console.error("An error occurred while adding items:", err);
        if (err.message.includes("Microsoft.SharePoint.SPDuplicateValuesFoundException")) {
          setError("This portal already exists, please go back and try again.");
        } else {
          setError("An unexpected error occurred: " + err.message);
        }
      });
  };

  const updateEngListRegularMatter = async (matterToUpdate, selectedPortalTypeShortHand) => {
    const item = await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).select("Portals_x0020_Created").get();
    if (item.Portals_x0020_Created === null) {
      await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).update({
        Portals_x0020_Created: selectedPortalTypeShortHand,
      });
    } else {
      await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).update({
        Portals_x0020_Created: item.Portals_x0020_Created + "," + selectedPortalTypeShortHand,
      });
    }
  };
  useEffect(() => {
    if (itemsStaged.length > 0) {
      setIsTeamAndPortalDisabled(true);
    } else {
      setIsTeamAndPortalDisabled(false);
    }
  }, [itemsStaged]);

  useEffect(() => {
    setEnableNextButton(checkItemsStagedForSiteOwner());
  }, [itemsStaged]);
  useEffect(() => {
    console.log("portalSelected::", portalSelected[0]);

    if (portalSelected.length > 0) {
      moveSelectedToStaged();
    }


  }, [portalSelected]);
  useLayoutEffect(() => {
    if (isBulkCreationOpen) {
      const clientInfo = new ClientInfoClass();
      const siteClientNumber = spContext._pageContext._web.serverRelativeUrl.split("/").pop(); // Renamed 'clientSiteNumber' to 'siteClientNumber'

      Promise.all([
        getMatterNumbersForClientSite(siteClientNumber)
      ]).then(([matterNumbersResponse]) => {
        const engagementListMatters = matterNumbersResponse.engagementListMatters;

        // Create a set of titles from engagementListMatters for quick lookup
        const engagementTitles = new Set(engagementListMatters.map(item => item.Title));


        // Sort the combined items alphabetically by title
        const sortedEngagementItems = engagementListMatters.sort((a, b) => a.newMatterEngagementName.localeCompare(b.newMatterEngagementName));

        setItems(sortedEngagementItems);
        setIsDataLoaded(sortedEngagementItems.length > 0);
      });

      let obj = new ClientInfoClass();
      // Load dropdown options for industry types, supplementals, and template types
      obj.GetIndustryTypes().then(data => setIndustryTypes(data.sort((a, b) => a.Title.localeCompare(b.Title))));
      obj.GetSupplemental().then(data => setSupplementals([{ Title: "N/A" }, ...data.sort((a, b) => a.Title.localeCompare(b.Title))]));
      obj.GetServiceTypes().then(data => setTemplateTypes(data.sort((a, b) => a.Title.localeCompare(b.Title))));
    }
  }, [isBulkCreationOpen]);


  const getYearsDropdown = (matterNumber: string, workYearsToExclude: string[]) => {
    const currentYear = new Date().getFullYear();
    const years = [];
    for (let i = currentYear - 5; i <= currentYear + 5; i++) {
      years.push(i.toString());
    }
    return matterNumber.endsWith("00") ? years.filter(year => !workYearsToExclude.includes(year)) : [currentYear.toString()];
  };

  const viewFieldsStaged: IViewField[] = [
    {
      name: "newMatterEngagementName",
      displayName: "Engagement Name",
      sorting: false,
      minWidth: 150,  // Increased minWidth
      maxWidth: 300,  // Increased maxWidth
      isResizable: true,
    },
    {
      name: "newMatterNumber",
      displayName: "Matter #",
      sorting: false,
      minWidth: 150,  // Increased minWidth
      maxWidth: 300,  // Increased maxWidth
      isResizable: true,
    },
    {
      name: "newMatterWorkYear",
      displayName: "Year",
      sorting: false,
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      render: (rowItem, index, column) => {
        const currentYear = new Date().getFullYear();
        let newMatterNumber = rowItem.newMatterNumber;
        if (rowItem.engagementNumberEndZero !== undefined) {
          newMatterNumber = rowItem.engagementNumberEndZero;
        }
        // Filter engagement items based on EngagementNumberEndZero
        const filteredEngagementItems = items.filter(item => item.engagementNumberEndZero === rowItem.engagementNumberEndZero);

        // Get work years to exclude
        const workYearsToExclude = filteredEngagementItems.map(item => item.WorkYear);

        const options: IDropdownOption[] = getYearsDropdown(newMatterNumber, workYearsToExclude).map((year) => ({
          key: year,
          text: year,
        }));
        const endsWithZeroZero = (number: string | undefined) => number && number.endsWith("00");
        return endsWithZeroZero(rowItem.newMatterNumber) || endsWithZeroZero(rowItem.engagementNumberEndZero) ? (
          <Dropdown
            selectedKey={rowItem.newMatterWorkYear || currentYear.toString()}
            onChange={(event, option) => {
              const selectedKey = option.key as string;
              const lastTwoDigits = selectedKey.slice(-2);

              const updatedItemsStaged = itemsStaged.map((item) => {
                if (item.ID === rowItem.ID) {
                  const updatedNewMatterNumber = item.newMatterNumber.slice(0, -2) + lastTwoDigits;
                  return {
                    ...item,
                    newMatterWorkYear: selectedKey,
                    newMatterNumber: updatedNewMatterNumber,
                    engagementNumberEndZero: item.newMatterNumber.endsWith("00") ? item.newMatterNumber : item.engagementNumberEndZero,
                  };
                }
                return item;
              });

              setItemsStaged(updatedItemsStaged);
            }}
            options={options}
            calloutProps={{ className: styles.wideDropdown }}
            className={styles.smallFont}
          />
        ) : (
          <span style={{ fontSize: '12px' }}>{rowItem.newMatterWorkYear}</span>
        );
      },
    },

    ...(team === "AUD" && portalType === "workflow" ? [
      {
        name: "industryType",
        displayName: "Industry Type",
        sorting: false,
        minWidth: 150,  // Increased minWidth
        maxWidth: 300,  // Increased maxWidth
        isResizable: true,
        render: (rowItem, index, column) => {
          const options: IDropdownOption[] = [{ key: "N/A", text: "N/A" }, ...industryTypes
            .filter(type => type.Title === "Assurance")
            .map((type) => ({
              key: type.IndustryType,
              text: type.IndustryType,
            }))];

          return (
            <Dropdown
              selectedKey={rowItem.industryType || "N/A"}
              onChange={(event, option) => {
                const updatedItemsStaged = itemsStaged.map((item) => {
                  if (item.ID === rowItem.ID) {
                    return {
                      ...item,
                      industryType: option.key as string,
                    };
                  }
                  return item;
                });
                setItemsStaged(updatedItemsStaged);
              }}
              options={options}
              calloutProps={{ className: styles.wideDropdown }}
              className={styles.smallFont}
            />
          );
        },
      },
      {
        name: "supplemental",
        displayName: "Supplemental",
        sorting: false,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        render: (rowItem, index, column) => {
          const distinctSupplementals = Array.from(new Set(supplementals.map(supp => supp.Title)));
          const options: IDropdownOption[] = distinctSupplementals.map((supp) => ({
            key: supp,
            text: supp,
          }));

          return (
            <Dropdown
              selectedKey={rowItem.supplemental || "N/A"}
              onChange={(event, option) => {
                const updatedItemsStaged = itemsStaged.map((item) => {
                  if (item.ID === rowItem.ID) {
                    return {
                      ...item,
                      supplemental: option.key as string,
                    };
                  }
                  return item;
                });
                setItemsStaged(updatedItemsStaged);
              }}
              options={options}
              calloutProps={{ className: styles.wideDropdown }}
              className={styles.smallFont}
            />
          );
        },
      },
    ] : []),
    ...(team === "TAX" && portalType === "workflow" ? [
      {
        name: "templateType",
        displayName: "Template Type",
        sorting: false,
        minWidth: 150,  // Increased minWidth
        maxWidth: 300,  // Increased maxWidth
        isResizable: true,
        render: (rowItem, index, column) => {
          const options: IDropdownOption[] = [{ key: "N/A", text: "N/A" }, ...templateTypes
            .filter(type => type.Title === "TAX")
            .map((type) => ({
              key: type.ServiceType,
              text: type.ServiceType,
            }))];

          return (
            <Dropdown
              selectedKey={rowItem.templateType || "N/A"}
              onChange={(event, option) => {
                const updatedItemsStaged = itemsStaged.map((item) => {
                  if (item.ID === rowItem.ID) {
                    return {
                      ...item,
                      templateType: option.key as string,
                    };
                  }
                  return item;
                });
                setItemsStaged(updatedItemsStaged);
              }}
              options={options}
              calloutProps={{ className: styles.wideDropdown }}
              className={styles.smallFont}
            />
          );
        },
      },
      {
        name: "industryType",
        displayName: "Industry Type",
        sorting: false,
        minWidth: 150,  // Increased minWidth
        maxWidth: 300,  // Increased maxWidth
        isResizable: true,
        render: (rowItem, index, column) => {
          const options: IDropdownOption[] = [{ key: "N/A", text: "N/A" }, ...industryTypes
            .filter(type => type.Title === "Tax")
            .map((type) => ({
              key: type.IndustryType,
              text: type.IndustryType,
            }))];

          return (
            <Dropdown
              selectedKey={rowItem.industryType || "N/A"}
              onChange={(event, option) => {
                const updatedItemsStaged = itemsStaged.map((item) => {
                  if (item.ID === rowItem.ID) {
                    return {
                      ...item,
                      industryType: option.key as string,
                    };
                  }
                  return item;
                });
                setItemsStaged(updatedItemsStaged);
              }}
              options={options}
              calloutProps={{ className: styles.wideDropdown }}
              className={styles.smallFont}
            />
          );
        },
      },
    ] : []),
    {
      name: "siteOwner",
      displayName: "Site Owner",
      sorting: false,
      minWidth: 180,
      maxWidth: 250,
      isResizable: true,
      render: (rowItem, index, column) => (
        <PeoplePicker
          context={spContext}
          showtooltip={false}
          required={true}
          onChange={(item) => validateSiteOwner(item, rowItem)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          ensureUser={true}
          personSelectionLimit={1}
          placeholder="Enter name or email"
          defaultSelectedUsers={rowItem["siteOwner.Email"] ? [rowItem["siteOwner.Email"]] : []}
        />
      ),
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Portal Expiration Date",
      sorting: false,
      minWidth: 150,
      maxWidth: 300,
      isResizable: false,
      render: (rowItem, index, column) => {
        const defaultDate = portalType === "workflow"
          ? new Date(new Date().setMonth(new Date().getMonth() + 18))
          : new Date(new Date().setMonth(new Date().getMonth() + 12));

        const stagedItem = itemsStaged.find(item => item.ID === rowItem.ID);
        const selectedDate = stagedItem && stagedItem.newMatterPortalExpirationDate !== ""
          ? new Date(stagedItem.newMatterPortalExpirationDate)
          : defaultDate;
        stagedItem.newMatterPortalExpirationDate = selectedDate.toISOString();
        return (
          <DatePicker
            allowTextInput={false}
            value={selectedDate}
            initialPickerDate={defaultDate}
            onSelectDate={(dateToSend) => onSelectDate(dateToSend, rowItem)}
            formatDate={onFormatDate}
            maxDate={createDate18MonthsFromNow()}
            className={styles.smallFont}
          />
        );
      },
    },
    {
      name: "",
      displayName: "",
      sorting: false,
      minWidth: 100,
      maxWidth: 100,
      isResizable: false,
      render: (rowItem, index, column) => (
        <Icon iconName='Delete' className={styles.trashCan} onClick={(ev) => unstageItem(ev, rowItem)} />
      ),
    },
  ];



  const filterItems = (selectedTeam: string, selectedPortalType: string) => {
    return items.filter(item => {
      const matchesTeam = item.team === selectedTeam;
      const hasPortalType = item.Portals_x0020_Created ? item.Portals_x0020_Created.includes(selectedPortalType === "workflow" ? "WF" : "FE") : false;
      return matchesTeam && !hasPortalType;
    }).sort((a, b) => a.newMatterEngagementName.localeCompare(b.newMatterEngagementName));
  };

  return (
    <>
      <Dialog
        hidden={!isBulkCreationOpen}
        onDismiss={resetState}
        minWidth={1200}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Bulk Subportal Creation",
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          className: styles.bulkCreation,
        }}
      >
        {isDataLoaded && !isConfirmationScreen && (
          <>

            <span className={styles.guidanceText}>
             Choose the type of portal
            </span>
            <div className={styles.choiceGroupContainer}>
              <ChoiceGroup
                className={styles.innerChoice}
                defaultSelectedKey={portalType}
                label="Portal Type"
                required={true}
                options={[
                  { key: "workflow", text: "Workflow" },
                  { key: "fileexchange", text: "File Exchange" }
                ]}
                onChange={onPortalTypeChange}
                disabled={isTeamAndPortalDisabled}  // Disable when rows are staged
              />
            </div>
            {portalType && (
              <>
                <span className={styles.guidanceText}>
                  Choose a team to see portals that are available for creation
                </span>
                <div className={styles.choiceGroupContainer}>
                  <ChoiceGroup
                    className={styles.innerChoice}
                    defaultSelectedKey={teamKey}
                    label="Team"
                    required={true}
                    options={[
                      { key: "assurance", text: "Assurance" },
                      { key: "tax", text: "Tax" },
                      { key: "advisory", text: "Advisory" }
                    ]}
                    onChange={onTeamChange}
                    disabled={isTeamAndPortalDisabled}  // Disable when rows are staged
                  />
                </div>
              </>
            )}
          </>
        )}

        {!isDataLoaded && !isConfirmationScreen && (
          <Spinner
            size={SpinnerSize.large}
            label="Loading Eligible Creation Portals...this could take some time depending on the amount of portals."
          />
        )}

        {isDataLoaded && team !== "" && portalType !== "" && !isConfirmationScreen && (
          <>
            <span className={styles.guidanceText}>
              Select engagements below for bulk creation.
            </span>
            <div className={styles.listViewPortsForCreation}>
              <ListView
                items={filterItems(team, portalType)}
                viewFields={viewFields}
                compact={true}
                selectionMode={SelectionMode.single}
                selection={(selectionItem) => setPortalSelected(selectionItem)}
                showFilter={false}
                key="engagementPortals"
              />
            </div>
            {itemsStaged.length > 0 && (
              <>
                <span className={styles.guidanceText}>
                  Enter a Site Owner and Expiration Date for each portal to creation.
                </span>
                <br />
                <span><i>
                  The portal will be available for future creation until the expiration date below. All files will be deleted from the portal 12 months from today's date.
                </i></span>
                <ListView
                  items={itemsStaged}
                  viewFields={viewFieldsStaged}
                  compact={true}
                  selectionMode={SelectionMode.none}
                  showFilter={false}
                  key="engagementPortalsStaged"
                />
              </>
            )}
          </>
        )}

        {isConfirmationScreen && (
          <>
            <span className={styles.guidanceText}>
              Selected engagements will be created over from previous year. No Permissions will be created over to the new portals.
            </span>
            <div className={styles.listViewPortsForCreation}>
              <ListView
                items={itemsStaged}
                viewFields={confirmationViewFields}
                compact={true}
                selectionMode={SelectionMode.none}
                showFilter={false}
                key="confirmationCreations"
              />
            </div>
            {isDataSubmitted && (
              <MessageBar
                messageBarType={MessageBarType.success}
                isMultiline={true}
                className={styles.successMsg}
              >
                Thank you. Your portals are in the process of being created. You will receive an email confirmation shortly when your portals are active. Please close this window.
              </MessageBar>
            )}
            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
                {error}
              </MessageBar>
            )}
          </>
        )}

        <DialogFooter>
          {isDataLoaded && team !== "" && portalType !== "" && (
            <div className={styles.dialogFooterButtonContainer}>
              <DefaultButton
                className={styles.defaultButton}
                onClick={resetState}
                text={isDataSubmitted ? "Close" : "Cancel"}
              />
              <div>
                {isConfirmationScreen && (
                  <DefaultButton
                    className={styles.defaultButton}
                    onClick={() => setIsConfirmationScreen(false)}
                    style={{ marginRight: "8px" }}
                    text="Back"
                    disabled={isDataSubmitted} // Disable when data is submitted
                  />
                )}
                {enableNextButton && !isConfirmationScreen && itemsStaged.length > 0 && (
                  <PrimaryButton
                    className={styles.primaryButton}
                    onClick={() => setIsConfirmationScreen(true)}
                    text="Next"
                  />
                )}
                {isConfirmationScreen && (
                  <PrimaryButton
                    className={styles.primaryButton}
                    onClick={submitPortalCreationData}
                    text="Create Portals"
                    disabled={isDataSubmitted}// Disable when data is submitted
                  />
                )}
              </div>
            </div>
          )}
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default BulkCreation;
