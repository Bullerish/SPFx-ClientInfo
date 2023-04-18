import * as React from "react";
import { useState, useEffect } from "react";
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
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Text } from "office-ui-fabric-react/lib/Text";
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";

import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import ConfirmDialog from './ConfirmDialog';
import StatusDialog from './StatusDialog';

// interface for event handler for dropdowns
interface IDropdownControlledState {
  jobFunctionItem?: { key: string | number | undefined };
  jobRoleItem?: { key: string | number | undefined };
}
// interface for all checkbox options: services, sectors, other interests
interface IServices {
  text: string;
  checked: boolean;
}
// set options for job function dropdown
const jobFunctionOptions = [
  { key: "Accounting", text: "Accounting" },
  { key: "Administration", text: "Administration" },
  { key: "Analytics", text: "Analytics" },
  { key: "Compliance", text: "Compliance" },
  { key: "Executive", text: "Executive" },
  { key: "Finance", text: "Finance" },
  { key: "Human Resources", text: "Human Resources" },
  { key: "IT", text: "IT" },
  { key: "Legal", text: "Legal" },
  { key: "Operations", text: "Operations" },
  { key: "Risk", text: "Risk" },
  { key: "Sales/Marketing", text: "Sales/Marketing" },
  { key: "Tax", text: "Tax" },
  { key: "Other", text: "Other" },
];
// set options for job role dropdown
const jobRoleOptions = [
  { key: "Board Member", text: "Board Member" },
  { key: "Csuite", text: "Csuite" },
  { key: "VP / Director", text: "VP / Director" },
  { key: "Manager", text: "Manager" },
  { key: "Staff", text: "Staff" },
  { key: "Student", text: "Student" },
  { key: "Other", text: "Other" },
];
// column props for
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};

const profileContainerProps: Partial<IStackProps> = {
  styles: { root: { width: 700, marginTop: 20 } },
};

const textAreasProps: Partial<IStackProps> = {
  styles: { root: { width: 650 } },
};

const ClientProfileInfo = ({spContext, isClientProfileInfoModalOpen, onClientProfileInfoModalHide,}): React.ReactElement => {
  // state for currentUser and Webs instances
  const [currentUser, setCurrentUser] = useState<ISiteUserInfo>(null);
  const [userItemData, setUserItemData] = useState([]);

  // states for confirmation and status dialogs
  const [isSubmissionSuccessful, setIsSubmissionSuccessful] = useState<boolean>(null);
  const [confirmDialogHidden, setConfirmDialogHidden] = useState<boolean>(true);
  const [statusDialogHidden, setStatusDialogHidden] = useState<boolean>(true);



  // flags for "!" icon and reminder toast
  const [isComplete, setIsComplete] = useState<boolean>(null);
  const [reminder, setReminder] = useState<boolean>(null);

  // Profile form tab states for inputs
  const [fullName, setFullName] = useState<string>('');
  const [title, setTitle] = useState<string>('');
  const [jobRoleItem, setJobRoleItem] = useState({
    key: undefined,
    text: "",
  });
  const [jobFunctionItem, setJobFunctionItem] = useState({
    key: undefined,
    text: "",
  });
  const [mailingAddress, setMailingAddress] = useState<string>('');
  const [esgInterests, setEsgInterests] = useState<string>('');
  const [boardPositions, setBoardPositions] = useState<string>('');
  const [passions, setPassions] = useState<string>('');

  // states for "Subscriptions" tab
  // for services checkboxes
  const [services, setServices] = useState<IServices[]>([
    { text: "Accounting/Assurance", checked: false },
    { text: "Advisory", checked: false },
    { text: "Private Client Services", checked: false },
    { text: "Solutions for C-Suite Executives", checked: false },
    { text: "Tax", checked: false },
  ]);
  // for sectors checkboxes
  const [sectors, setSectors] = useState<IServices[]>([
    { text: "Affordable Housing", checked: false },
    { text: "Construction", checked: false },
    { text: "Commercial Real Estate", checked: false },
    { text: "Financial Services", checked: false },
    { text: "Cannabis", checked: false },
    { text: "Healthcare", checked: false },
    { text: "Life Sciences", checked: false },
    { text: "Manufacturing and Distribution", checked: false },
    { text: "Not for Profit and Education", checked: false },
    { text: "Private Equity/Other Financial Sponsors", checked: false },
    { text: "Renewable Energy", checked: false },
    { text: "Retail", checked: false },
    { text: "Government Contracting", checked: false },
    { text: "CanGovernment - Audit/Accountingnabis", checked: false },
    { text: "Government - Compliance and Monitoring", checked: false },
    { text: "Government - Emergency Management", checked: false },
  ]);
  // for other interests checkboxes
  const [otherInterests, setOtherInterests] = useState<IServices[]>([
    { text: "Alumni Events", checked: false },
    { text: "CPE Offerings", checked: false },
    { text: "Executive Women's Forum Events", checked: false },
  ]);

  // state for "Contacts" tab
  const [contacts, setContacts] = useState<string>('');

  // list name of the Client Profile List to save/retrieve user item data in /sites/clientportal
  const clientProfileListName = 'ClientProfileList';

  // this useEffect will run only once after initially page load/component mount
  useEffect(() => {
    const siteWebVal = Web(GlobalValues.SiteURL);

    const getCurrentUser = async () => {
      const userInfo = await siteWebVal.currentUser();
      setCurrentUser(userInfo);
    }

    getCurrentUser();

  }, []);

  // check list for current logged in user list item. Set the usetItemData state variable with results
  useEffect(() => {
    console.log(currentUser);

    // check to see if user has existing item, if so bring it back
    const checkSetUserItem = async () => {
      const hubWeb = Web(GlobalValues.HubSiteURL);
      const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.filter(`UserLoginName eq '${currentUser.LoginName}'`).get();
      console.log('logging userItem:: ', userItem);

      // TODO: will have to update form complete/reminder criteria based on marketing/courtney feedback
      if (!userItem.length) {
        setIsComplete(false);
        setReminder(true);
      }

      setUserItemData(userItem);
    };

    checkSetUserItem();

  }, [currentUser]);

  const stateLogger = () => {
    console.log('logging all state::');

  };


  // set services checkboxes and state
  const onServicesChange = (text: string, value) => {
    let servicesTempArr = services;

    servicesTempArr.forEach((e) => {
      if (text === e.text && value) {
        console.log("inside foreach logging val:: ", e);
        e.checked = true;
      }

      if (text === e.text && !value) {
        e.checked = false;
      }
    });

    setServices([...servicesTempArr]);
  };

   // set services checkboxes and state
   const onSectorsChange = (text: string, value) => {
    let sectorsTempArr = sectors;

    sectorsTempArr.forEach((e) => {
      if (text === e.text && value) {
        e.checked = true;
      }

      if (text === e.text && !value) {
        e.checked = false;
      }
    });

    setSectors([...sectorsTempArr]);
  };

  // set other interests checkboxes and state
  const onOtherInterestsChange = (text: string, value) => {
    let otherInterestsTempArr = otherInterests;

    otherInterestsTempArr.forEach((e) => {
      if (text === e.text && value) {
        e.checked = true;
      }

      if (text === e.text && !value) {
        e.checked = false;
      }
    });

    setOtherInterests([...otherInterestsTempArr]);
  };

  // event handler to hide confirmation dialog
  const onSetConfirmDialogHidden = () => {
    setConfirmDialogHidden(true);
  };

  // event handler to hide status dialog
  const onSetStatusDialogHidden = () => {
    setStatusDialogHidden(true);
  };

  return (
    <div>
      <Dialog
        hidden={!isClientProfileInfoModalOpen}
        onDismiss={onClientProfileInfoModalHide}
        minWidth={750}
        dialogContentProps={{
          type: DialogType.normal,
          title: "My Profile Information",
          showCloseButton: true,
        }}
        modalProps={{
          isBlocking: true,
          // styles: { main: { maxHeight: 700, overflowY: 'scroll' } },
          className: styles.clientProfileInfo,
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <Pivot>
          {/* PROFILE TAB */}
          <PivotItem
            headerText="Profile"
            headerButtonProps={{
              "data-order": 1,
              "data-title": "Profile Title",
            }}
          >
            {/* parent stack wrapper */}
            <Stack gap={20} horizontalAlign="center" {...profileContainerProps}>
              {/* presentation/alignment stack for form */}
              <Stack
                horizontal
                horizontalAlign="center"
                tokens={{ childrenGap: 50 }}
                styles={{ root: { width: 650 } }}
              >

                <Stack {...columnProps}>
                  <TextField label="Full Name" value={fullName} onChange={(ev, newValue) => setFullName(newValue)} />

                  <Dropdown
                    label="Job Role"
                    selectedKey={
                      jobRoleItem ? jobRoleItem.key : undefined
                    }
                    onChange={(ev, item) => setJobRoleItem(item)}
                    placeholder="Select an option"
                    options={jobRoleOptions}
                  />
                </Stack>

                <Stack {...columnProps}>
                  <TextField label="Title" value={title} onChange={(ev, newValue) => setTitle(newValue)} />

                  <Dropdown
                    label="Job Function"
                    selectedKey={
                      jobFunctionItem ? jobFunctionItem.key : undefined
                    }
                    onChange={(ev, item) => setJobFunctionItem(item)}
                    placeholder="Select an option"
                    options={jobFunctionOptions}
                  />
                </Stack>
              </Stack>

              <Stack gap={20} {...textAreasProps}>
                <TextField label="Mailing Address" multiline rows={4} value={mailingAddress} onChange={(ev, newValue) => setMailingAddress(newValue)} />
                <TextField
                  label="What are your ESG Interests?"
                  multiline
                  rows={4}
                  value={esgInterests}
                  onChange={(ev, newValue) => setEsgInterests(newValue)}
                />
                <TextField
                  label="Do you hold any board positions, if so which boards?"
                  multiline
                  rows={4}
                  value={boardPositions}
                  onChange={(ev, newValue) => setBoardPositions(newValue)}
                />
                <TextField
                  label="What are you passionate about?"
                  multiline
                  rows={4}
                  value={passions}
                  onChange={(ev, newValue) => setPassions(newValue)}
                />
              </Stack>
            </Stack>

          </PivotItem>
          {/* SUBSCRIPTIONS TAB */}
          {/* TODO: need to implement the state handling for all the checkboxes */}
          <PivotItem headerText="Subscriptions" className={styles.marginTabsTop}>

            <div>

              <div className={styles.subscriptionGuidanceText}>
                <Text variant="medium">Please choose the areas in which you would like to subscribe to learn more:</Text>
              </div>

              <Text variant="mediumPlus">SERVICES</Text>

              <div className={styles.checkboxFlexStyles}>

                {services.map((e) => (
                  <Checkbox
                  label={e.text}
                  checked={e.checked}
                  onChange={(ev, value) => onServicesChange(e.text, value)}
                  />
                ))}

              </div>

            </div>

            <div>
              <Text variant="mediumPlus">SECTORS</Text>

              <div className={styles.checkboxFlexStyles}>

                {sectors.map((e) => (
                  <Checkbox
                  label={e.text}
                  checked={e.checked}
                  onChange={(ev, value) => onSectorsChange(e.text, value)}
                  />
                ))}

              </div>
            </div>

            <div>
              <Text variant="mediumPlus">OTHER INTERESTS</Text>

              <div className={styles.checkboxFlexStyles}>

                {otherInterests.map((e) => (
                  <Checkbox
                  label={e.text}
                  checked={e.checked}
                  onChange={(ev, value) => onOtherInterestsChange(e.text, value)}
                  />
                ))}
              </div>
            </div>

          </PivotItem>
          {/* CONTACTS TAB */}
          <PivotItem headerText="Contacts" className={styles.marginTabsTop}>

            <div>
              <TextField
                label="Please share the names and emails of your team members who will be working with CohnReznick"
                multiline
                rows={6}
                value={fullName}
                onChange={(ev, newValue) => setContacts(newValue)}
              />
            </div>

          </PivotItem>
        </Pivot>
        <DialogFooter>
          <PrimaryButton
            className={styles.primaryButton}
            // onClick={ensureAlertsListExists}
            text="Save"
          />
          <DefaultButton
            className={styles.defaultButton}
            // onClick={() => setIsConfirmationHidden(true)}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
      <ConfirmDialog confirmDialogHidden={confirmDialogHidden} onSetConfirmDialogHidden={onSetConfirmDialogHidden} />
      <StatusDialog isSubmissionSuccessful={isSubmissionSuccessful} statusDialogHidden={statusDialogHidden} onSetStatusDialogHidden={onSetStatusDialogHidden} />
    </div>
  );
};

export default ClientProfileInfo;
