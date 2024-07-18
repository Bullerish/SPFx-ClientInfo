import * as React from 'react';
import { useState, useEffect, useLayoutEffect } from "react";
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

const BulkCreation = ({
  spContext,
  isBulkCreationOpen,
  onBulkCreationModalHide,
}): React.ReactElement => {
  const [team, setTeam] = useState<string>("");
  const [portalType, setPortalType] = useState<string>("");
  const [isDataLoaded, setIsDataLoaded] = useState<boolean>(false);
  const [items, setItems] = useState<MatterAndCreationData[]>([]);
  const [itemsStaged, setItemsStaged] = useState<MatterAndCreationData[]>([]);
  const [portalSelected, setPortalSelected] = useState([]);
  const [enableNextButton, setEnableNextButton] = useState<boolean>(false);
  const [isConfirmationScreen, setIsConfirmationScreen] = useState<boolean>(false);
  const [isDataSubmitted, setIsDataSubmitted] = useState<boolean>(false);

  const clientSiteAbsoluteUrl = spContext._pageContext._web.absoluteUrl;
  const clientSiteServerRelativeUrl = spContext._pageContext._web.serverRelativeUrl;
  const relativeUrlArr = clientSiteServerRelativeUrl.split("/");
  const clientSiteNumber = relativeUrlArr[relativeUrlArr.length - 1];
  const hubSite = Web(GlobalValues.HubSiteURL);

  const resetState = () => {
    onBulkCreationModalHide(false);
    setTeam("");
    setPortalType("");
    setIsDataLoaded(false);
    setItems([]);
    setItemsStaged([]);
    setPortalSelected([]);
    setEnableNextButton(false);
    setIsConfirmationScreen(false);
    setIsDataSubmitted(false);
  };

  const onTeamChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    setTeam(option.key);
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
          newMatterPortalExpirationDate: date.toString(),
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
    portalSelected.forEach((selectedItem) => {
      const newItems = items.filter((item) => item.ID !== selectedItem.ID);
      setItems(newItems);
      const isAlreadyStaged = itemsStaged.some((item) => item.ID === selectedItem.ID);
      if (!isAlreadyStaged) {
        setItemsStaged((prevItemsStaged) => [...prevItemsStaged, selectedItem]);
      }
    });
  };

  const unstageItem = (ev, itemRowToRemove) => {
    const newItemsStaged = itemsStaged.filter((item) => item.ID !== itemRowToRemove.ID);
    setItemsStaged(newItemsStaged);
    const isAlreadyInItems = items.some((item) => item.ID === itemRowToRemove.ID);
    if (!isAlreadyInItems) {
      itemRowToRemove.siteOwner = "";
      setItems((prevItems) => [...prevItems, itemRowToRemove]);
    }
  };

  const checkItemsStagedForSiteOwner = () => {
    return itemsStaged.every((item) => item.siteOwner && item.siteOwner !== "");
  };

  const submitPortalCreationData = () => {
    let mattersToUpdatePC = [];
    Promise.all(itemsStaged.map(stagedItem => {
      if (stagedItem.engagementNumberEndZero === "") {
        mattersToUpdatePC.push(stagedItem.engListID);
      }
      const itemData = {
        Title: stagedItem.newMatterNumber,
        EngagementName: stagedItem.newMatterEngagementName,
        ClientNumber: stagedItem.clientNumber,
        EngagementNumberEndZero: stagedItem.engagementNumberEndZero,
        WorkYear: stagedItem.newMatterWorkYear,
        Team: stagedItem.team,
        PortalType: stagedItem.portalType,
        SiteUrl: {
          __metadata: { type: "SP.FieldUrlValue" },
          Description: stagedItem.newMatterSiteUrl,
          Url: stagedItem.newMatterSiteUrl,
        },
        CreationUrl: {
          __metadata: { type: "SP.FieldUrlValue" },
          Description: stagedItem.creationMatterSiteUrl,
          Url: stagedItem.creationMatterSiteUrl,
        },
        Creation: stagedItem.creation,
        PortalId: stagedItem.newMatterPortalId,
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
          updateEngListRegularMatter(matterToUpdate);
        });
      }
    })
    .catch((error) => {
      console.error("An error occurred while adding items:", error);
    });
  };

  const updateEngListRegularMatter = async (matterToUpdate) => {
    const item = await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).select("Portals_x0020_Created").get();
    if (item.Portals_x0020_Created === null) {
      await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).update({
        Portals_x0020_Created: "WF",
      });
    } else {
      await hubSite.lists.getByTitle("Engagement List").items.getById(matterToUpdate).update({
        Portals_x0020_Created: item.Portals_x0020_Created + ",WF",
      });
    }
  };

  useEffect(() => {
    setEnableNextButton(checkItemsStagedForSiteOwner());
  }, [itemsStaged]);

  useLayoutEffect(() => {
    if (isBulkCreationOpen) {
      getMatterNumbersForClientSite(clientSiteNumber).then((response) => {
        setItems(response.engagementListMatters);
        setIsDataLoaded(response.engagementListMatters.length > 0);
      });
    }
  }, [isBulkCreationOpen]);

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
    {
      name: "templateType",
      displayName: "Template Type",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
  ];

  const viewFieldsStaged: IViewField[] = [
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
    {
      name: "templateType",
      displayName: "Template Type",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
    {
      name: "SiteOwner",
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
      minWidth: 125,
      maxWidth: 250,
      isResizable: false,
      render: (rowItem, index, column) => (
        <DatePicker
          allowTextInput={false}
          value={new Date(rowItem.newMatterPortalExpirationDate)}
          initialPickerDate={new Date()}
          onSelectDate={(dateToSend) => onSelectDate(dateToSend, rowItem)}
          formatDate={onFormatDate}
          maxDate={createDate18MonthsFromNow()}
        />
      ),
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

  const confirmationViewFields: IViewField[] = [
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
    {
      name: "templateType",
      displayName: "Template Type",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
    {
      name: "siteOwner",
      displayName: "Site Owner",
      sorting: false,
      minWidth: 100,
      maxWidth: 100,
      isResizable: false,
      render: (rowItem, index, column) => (
        <span>{rowItem["siteOwner.Title"]}</span>
      ),
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Expiration Date",
      sorting: false,
      minWidth: 100,
      maxWidth: 100,
      isResizable: false,
      render: (rowItem, index, column) => (
        <span>{onFormatDate(new Date(rowItem.newMatterPortalExpirationDate)).toString()}</span>
      ),
    },
  ];

  const filterItems = (selectedTeam: string, selectedPortalType: string) => {
    return items.filter(item => {
      const matchesTeam = item.team.toLowerCase() === selectedTeam.toLowerCase();
      const hasPortalType = item.Portals_x0020_Created ? item.Portals_x0020_Created.includes(selectedPortalType === "workflow" ? "WF" : "FE") : false;
      return matchesTeam && !hasPortalType;
    });
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
              Choose a team to see WF portals that are available for creation
            </span>
            <div className={styles.choiceGroupContainer}>
              <ChoiceGroup
                className={styles.innerChoice}
                defaultSelectedKey={team}
                label="Team"
                required={true}
                options={[
                  { key: "assurance", text: "Assurance" },
                  { key: "tax", text: "Tax" },
                  { key: "advisory", text: "Advisory"}
                ]}
                onChange={onTeamChange}
              />
            </div>
            {team && (
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
          </>
        )}

        <DialogFooter>
          {isDataLoaded && team !== "" && portalType !== "" && (
            <div className={styles.dialogFooterButtonContainer}>
              <DefaultButton
                className={styles.defaultButton}
                onClick={resetState}
                text="Cancel"
              />
              <div>
                {isConfirmationScreen && (
                  <DefaultButton
                    className={styles.defaultButton}
                    onClick={() => setIsConfirmationScreen(false)}
                    style={{ marginRight: "8px" }}
                    text="Back"
                  />
                )}
                {enableNextButton && !isConfirmationScreen && (
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
