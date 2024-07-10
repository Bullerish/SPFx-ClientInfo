import * as React from 'react';
import { useState, useEffect, useLayoutEffect } from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import {
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react/lib/MessageBar";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Text } from "office-ui-fabric-react/lib/Text";
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";
import { IItemAddResult } from "@pnp/sp/items";
import {
  DatePicker,
  DayOfWeek,
  IDatePickerStrings,
} from "office-ui-fabric-react/lib/DatePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  ListView,
  IViewField,
  SelectionMode,
  GroupOrder,
  IGrouping,
} from "@pnp/spfx-controls-react/lib/ListView";
import { getMatterNumbersForClientSite } from './creationLogic';
import { MatterAndCreationData } from './creationLogic';
import { set } from '@microsoft/sp-lodash-subset';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { createDate18MonthsFromNow } from './creationLogic';





export interface IDatePickerFormatExampleState {
  firstDayOfWeek?: DayOfWeek;
  value?: Date | null;
}

const BulkCreation = ({
  spContext,
  isBulkCreationOpen,
  onBulkCreationModalHide,
}): React.ReactElement => {
  // all state variables
  const [team, setTeam] = useState<string>("");
  const [isDataLoaded, setIsDataLoaded] = useState<boolean>(false);
  const [taxCreationData, setTaxCreationData] = useState<MatterAndCreationData[]>([]);
  const [AudCreationData, setAudCreationData] = useState<
    MatterAndCreationData[]
  >([]);

  const [items, setItems] = useState<MatterAndCreationData[]>([]);

  const [itemsStaged, setItemsStaged] = useState<MatterAndCreationData[]>([]);

  const [portalSelected, setPortalSelected] = useState([]);
  const [dateSelections, setDateSelections] = useState({});
  const [enableNextButton, setEnableNextButton] = useState<boolean>(false);
  const [isConfirmationScreen, setIsConfirmationScreen] = useState<boolean>(false);
  const [isDataSubmitted, setIsDataSubmitted] = useState<boolean>(false);

  // store the current site's absolute URL (should be a client site URL)
  const clientSiteAbsoluteUrl = spContext._pageContext._web.absoluteUrl;
  const hubSite = Web(GlobalValues.HubSiteURL);

  const clientSiteServerRelativeUrl =
    spContext._pageContext._web.serverRelativeUrl;
  const relativeUrlArr = clientSiteServerRelativeUrl.split("/");
  const clientSiteNumber = relativeUrlArr[relativeUrlArr.length - 1];

  // function to reset all state variables back to their initial values
  const resetState = () => {
    onBulkCreationModalHide(false);
    setTeam("");
    setIsDataLoaded(false);
    setTaxCreationData([]);
    setAudCreationData([]);
    setItems([]);
    setItemsStaged([]);
    setPortalSelected([]);
    setDateSelections({});
    setEnableNextButton(false);
    setIsConfirmationScreen(false);
    setIsDataSubmitted(false);
  };

  // onChange event handler to capture the selected team value
  const onTeamChange = (
    ev: React.FormEvent<HTMLInputElement>,
    option: any
  ): void => {
    console.log("onTeamChange fired::");
    // console.log(option.key);
    setTeam(option.key);
  };

  // function to format the date for the DatePicker component
  const onFormatDate = (date: Date): string => {
    // return (
    //   date.getMonth() +
    //   1 +
    //   "/" +
    //   date.getDate() +
    //   "/" +
    //   (date.getFullYear() % 100)
    // );
    return (
      date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear()
    );
  };

  // function to capture the selected date from the DatePicker component
  const onSelectDate = (
    date: Date | null | undefined,
    rowItemToUpdate: MatterAndCreationData
  ): void => {
    console.log('logging rowItemToUpdate with new::', rowItemToUpdate);
    // let tempItemsStaged = itemsStaged;


    const updatedItemsStaged = itemsStaged.map((item) => {
      if (item.ID === rowItemToUpdate.ID) {
        return {
          ...item,
          newMatterPortalExpirationDate: date.toString(),
        };
      }
      return item;
    });

    console.log('logging rowItem after date update::', updatedItemsStaged);

    setItemsStaged(updatedItemsStaged);

  };

  const getPeoplePickerItems = (itemsArr: any[], itemRow: MatterAndCreationData) => {
    const currSite = Web(GlobalValues.HubSiteURL);
    let getSelectedUsers = [];
    let getusersEmails = [];

    console.log("logging itemsArr::", itemsArr);


    for (let item in itemsArr) {
      getSelectedUsers.push(itemsArr[item].text);
      getusersEmails.push(itemsArr[item].secondaryText);
    }
    itemsArr.forEach((e) => {
      currSite.siteUsers
        .getByLoginName(e.loginName)
        .get()
        .then((user) => {
          // this.setState({
          //   addusers: getSelectedUsers,
          //   addusersID: user.Id,
          //   emailaddress: getusersEmails,
          // });
          console.log("logging user info:: ", user);
          console.log('logging itemRow::', itemRow);

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
    console.log('logging itemsSiteOwner::', itemsSiteOwner);
    console.log('logging rowItemToUpdate::', rowItemToUpdate);
    let tempItemsStaged = itemsStaged;
    // show error message if this is a guest user
    if (itemsSiteOwner.length > 0) {
      let userEmail = itemsSiteOwner[0].secondaryText.toLowerCase();
      if (
        userEmail.indexOf("cohnreznick.com") == -1 &&
        userEmail.indexOf("cohnreznickdev") == -1
      ) {
        // this is a guest user, do not validate
        // this.setState({ addusers: [] });
      } else {
        getPeoplePickerItems(itemsSiteOwner, rowItemToUpdate);
      }
    } else {
      // this.setState({ addusers: [] });

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
    // Step 1: Iterate over portalSelected to handle each selected item
    portalSelected.forEach((selectedItem) => {
      // Step 2: Remove the selected item from 'items'
      const newItems = items.filter((item) => item.ID !== selectedItem.ID);

      // Update the 'items' state without the selected item
      setItems(newItems);

      // Step 3: Add the selected item to 'itemsStaged'
      // Check if the item is already in 'itemsStaged' to avoid duplicates
      const isAlreadyStaged = itemsStaged.some(
        (item) => item.ID === selectedItem.ID
      );
      if (!isAlreadyStaged) {
        setItemsStaged((prevItemsStaged) => [...prevItemsStaged, selectedItem]);
      }
    });
  };

  const unstageItem = (ev, itemRowToRemove) => {
    console.log('logging itemRowToRemove::', itemRowToRemove);
    // Step 1: Remove the selected item from 'itemsStaged'
    const newItemsStaged = itemsStaged.filter(
      (item) => item.ID !== itemRowToRemove.ID
    );

    setItemsStaged(newItemsStaged);

    // Step 2: Add the selected item to 'items'
    // Check if the item is already in 'items' to avoid duplicates
    const isAlreadyInItems = items.some((item) => item.ID === itemRowToRemove.ID);
    if (!isAlreadyInItems) {

      itemRowToRemove.siteOwner = "";

      setItems((prevItems) => [...prevItems, itemRowToRemove]);
    }

  };

  // function to check if all items in the itemsStaged array contain a site owner
  const checkItemsStagedForSiteOwner = () => {
    let allItemsHaveSiteOwner = true;
    itemsStaged.forEach((item) => {
      if (item.siteOwner === "" || item.siteOwner === null) {
        allItemsHaveSiteOwner = false;
      }
    });
    return allItemsHaveSiteOwner;
  };

  const submitPortalCreationData = () => {
    // Step 1: Iterate over itemsStaged to handle each staged item
    Promise.all(itemsStaged.map(stagedItem => {
      // console.log("logging stagedItem::", stagedItem);
      // Step 2: Prepare the item data to be submitted
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
        isNotificationEmail: true,
      };

      // Step 3: Submit the item data to the list
      return hubSite.lists
        .getByTitle("Engagement Portal List")
        .items.add(itemData);
        // .then((result: IItemAddResult) => {
        //   console.log(`Item with ID: ${result.data.ID} added successfully`);
        // });
    }))
    .then((results) => {
      console.log('setting isDataSubmitted to true::');
      setIsDataSubmitted(true);

    })
    .catch((error) => {
      console.error("An error occurred while adding items:", error);
    });



  };

  useEffect(() => {
    console.log("items::", items);
  }, [items]);

  useEffect(() => {
    console.log("itemsStaged::", itemsStaged);
    if (itemsStaged.length > 0) {
      setEnableNextButton(checkItemsStagedForSiteOwner());
    }
  }, [itemsStaged]);


  useEffect(() => {
    console.log("team selected::", team);

    setItemsStaged([]);

    if (team === "tax") {
      console.log("logging taxCreationData::", taxCreationData);
      setItems(taxCreationData);
    } else if (team === "assurance") {
      console.log("logging audCreationData::", AudCreationData);
      setItems(AudCreationData);
    }

    // getMatterNumbersForClientSite();
  }, [team]);

  useEffect(() => {
    console.log("portalSelected::", portalSelected[0]);

    if (portalSelected.length > 0) {
      moveSelectedToStaged();
    }


  }, [portalSelected]);

  useEffect(() => {
    console.log("dateSelections::", dateSelections);
  }, [dateSelections]);

  useEffect(() => {
    console.log("enableNextButton::", enableNextButton);
  }, [enableNextButton]);

  useLayoutEffect(() => {

    if (isBulkCreationOpen) {
      console.log(
        "useLayoutEffect fired, calling getMatterNumbersForClientSite::"
      );

      getMatterNumbersForClientSite(clientSiteNumber).then((response) => {
        console.log("logging response from getMatterNumbersForClientSite::", response);

        setAudCreationData(response.audMatters);
        setTaxCreationData(response.taxMatters);

        if (response.audMatters.length > 0 || response.taxMatters.length > 0) {
          setIsDataLoaded(true);
        }


        // setItems(response);
        // setItemsStaged(response);
      });

      // console.log('logging taxAndAudMatters from BulkCreation useLayoutEffect::', taxAndAudMatters);

    }
  }, [isBulkCreationOpen]);

  // define columns/viewfields so the ListView component knows what to render
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

  // define columns/viewfieldsStaged so the ListView component knows what to render (is for staged portals)
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
      render: (rowItem, index, column) => {

        console.log("logging rowItem from viewFieldsStaged::", rowItem);

        // console.log('logging rowItem siteOwner Email::', (rowItem as ISiteUserInfo)["siteOwner.Email"]);

        return (
          <div>
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
          </div>
        );
      },
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Expiration Date",
      sorting: false,
      minWidth: 125,
      maxWidth: 250,
      isResizable: false,
      render: (rowItem, index, column) => {
        // console.log("rowItem::", rowItem);

        // const onDateChange = (date: Date | null | undefined): void => {
        //   setDateSelections((prevSelections) => ({
        //     ...prevSelections,
        //     [rowItem.ID]: date, // Use a unique identifier from your row data
        //   }));
        // };

        return (
          <div>
            <DatePicker
              // label="DateTime Picker - 24h"
              // dateConvention={DateConvention.DateTime}
              // timeConvention={TimeConvention.Hours24}
              allowTextInput={false}
              value={new Date(rowItem.newMatterPortalExpirationDate)}
              initialPickerDate={new Date()}
              onSelectDate={(dateToSend) => onSelectDate(dateToSend, rowItem)}
              formatDate={onFormatDate}
            />
          </div>
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
      render: (rowItem, index, column) => {


        return (
          <div>
            <Icon iconName='Delete' className={styles.trashCan} onClick={(ev) => unstageItem(ev, rowItem)} />
          </div>
        );
      },
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
      render: (rowItem, index, column) => {
        // console.log("logging rowItem from confirmationViewFields::", rowItem);

        return (
          <div>
            <span>{rowItem["siteOwner.Title"]}</span>
          </div>
        );
      },
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Expiration Date",
      sorting: false,
      minWidth: 100,
      maxWidth: 100,
      isResizable: false,
      render: (rowItem, index, column) => {
        // console.log("logging rowItem from confirmationViewFields::", rowItem);

        return (
          <div>
            <span>
              {onFormatDate(
                new Date(rowItem.newMatterPortalExpirationDate)
              ).toString()}
            </span>
          </div>
        );
      },
    },
  ];

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
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
          className: styles.bulkCreation,
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        {/* Team select choice group */}
        {isDataLoaded && !isConfirmationScreen && (
          <>
            <span className={styles.guidanceText}>
              Choose a team to see WF portals that are available for creation
            </span>
            <div className={styles.choiceGroupContainer}>
              <ChoiceGroup
                className={styles.innerChoice}
                // defaultSelectedKey="B"
                label="Team"
                required={true}
                options={[
                  {
                    key: "assurance",
                    text: "Assurance",
                  },
                  {
                    key: "tax",
                    text: "Tax",
                  },
                ]}
                onChange={onTeamChange}
              />
            </div>
          </>
        )}

        {!isDataLoaded && !isConfirmationScreen && (
          <Spinner
            size={SpinnerSize.large}
            label="Loading Eligible Creation Portals...this could take some time depending on the amount of portals."
          />
        )}

        {/* ListView component to display list of portals */}
        {isDataLoaded && team !== "" && !isConfirmationScreen && (
          <>
            <span className={styles.guidanceText}>
              Select engagements below to bulk creation. No permissions will be
              rolled over to the new portals.
            </span>

            {/* ListView component to hold portals available for creation */}
            <div className={styles.listViewPortsForCreation}>
              <ListView
                items={items}
                viewFields={viewFields}
                // iconFieldName="FileRef"
                compact={true}
                selectionMode={SelectionMode.single}
                selection={(selectionItem) => setPortalSelected(selectionItem)}
                // defaultSelection={defaultSelectedFromScreen2}
                showFilter={false}
                key="engagementPortals"
              />
            </div>

            <span className={styles.guidanceText}>
              Enter a Site Owner and Expiration Date for each portal to
              creation. No permissions will be rolled over to the new portals.
            </span>

            {/* ListView component to hold staged portals */}
            <ListView
              items={itemsStaged}
              viewFields={viewFieldsStaged}
              // iconFieldName="FileRef"
              compact={true}
              selectionMode={SelectionMode.none}
              // selection={(selectionItem) => setPortalSelected(selectionItem)}
              // defaultSelection={defaultSelectedFromScreen2}
              showFilter={false}
              key="engagementPortalsStaged"
            />
          </>
        )}

        {isConfirmationScreen && (
          <>
            <span className={styles.guidanceText}>
              Selected engagements will be rolled over from previous year. No
              Permissions will be rolled over to the new portals.
            </span>

            {/* ListView component to hold portals available for creation */}
            <div className={styles.listViewPortsForCreation}>
              <ListView
                items={itemsStaged}
                viewFields={confirmationViewFields}
                // iconFieldName="FileRef"
                compact={true}
                selectionMode={SelectionMode.none}
                // selection={(selectionItem) => setPortalSelected(selectionItem)}
                // defaultSelection={defaultSelectedFromScreen2}
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
                Thank you. Your portals are in the process of being created. You
                will receive an email confirmation shortly when your portals are
                active. Please close this window.
              </MessageBar>
            )}
          </>
        )}

        {/* Dialog footer to hold buttons */}
        <DialogFooter>
          {isDataLoaded && team !== "" && (
            <>
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
                      // disabled={!enableNextButton}
                    />
                  )}

                  {isConfirmationScreen && (
                    <PrimaryButton
                      className={styles.primaryButton}
                      onClick={submitPortalCreationData}
                      text="Create Portals"
                      // disabled={!enableNextButton}
                    />
                  )}
                </div>
              </div>
            </>
          )}
        </DialogFooter>
      </Dialog>
      {/* confirm dialog component. Modal/dialog window will open */}
      {/* <ConfirmDialog
        confirmDialogHidden={confirmDialogHidden}
        onSetConfirmDialogHidden={onSetConfirmDialogHidden}
        onConfirmSubmission={submitUserProfileInfo}
      /> */}
    </>
  );
};

export default BulkCreation;
