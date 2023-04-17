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
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Text } from "office-ui-fabric-react/lib/Text";
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteUser } from "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import styles from "../ClientInfoWebpart.module.scss";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { setBaseUrl } from "office-ui-fabric-react";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";


interface IDropdownControlledState {
  jobFunctionItem?: { key: string | number | undefined };
  jobRoleItem?: { key: string | number | undefined };
}

interface IServices {
  text: string;
  checked: boolean;
}

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

const jobRoleOptions = [
  { key: "Board Member", text: "Board Member" },
  { key: "Csuite", text: "Csuite" },
  { key: "VP / Director", text: "VP / Director" },
  { key: "Manager", text: "Manager" },
  { key: "Staff", text: "Staff" },
  { key: "Student", text: "Student" },
  { key: "Other", text: "Other" },
];

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

const ClientProfileInfo = ({
  spContext,
  isClientProfileInfoModalOpen,
  onClientProfileInfoModalHide,
}): React.ReactElement => {
  const [jobFunctionItem, setJobFunctionItem] = useState({
    key: undefined,
    text: "",
  });

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

  useEffect(() => {
    console.log("triggering useEffect services");
    console.log(services);
  }, [services]);

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

            <Stack gap={20} horizontalAlign="center" {...profileContainerProps}>
              <Stack
                horizontal
                horizontalAlign="center"
                tokens={{ childrenGap: 50 }}
                styles={{ root: { width: 650 } }}
              >
                <Stack {...columnProps}>
                  <TextField label="Full Name" />

                  <Dropdown
                    label="Job Role"
                    selectedKey={
                      jobFunctionItem ? jobFunctionItem.key : undefined
                    }
                    // onChange={this._onChange}
                    placeholder="Select an option"
                    options={jobRoleOptions}
                    styles={{ dropdown: { width: 300 } }}
                  />
                </Stack>

                <Stack {...columnProps}>
                  <TextField label="Title" />

                  <Dropdown
                    label="Job Function"
                    selectedKey={
                      jobFunctionItem ? jobFunctionItem.key : undefined
                    }
                    // onChange={this._onChange}
                    placeholder="Select an option"
                    options={jobFunctionOptions}
                    styles={{ dropdown: { width: 300 } }}
                  />
                </Stack>
              </Stack>

              <Stack gap={20} {...textAreasProps}>
                <TextField label="Mailing Address" multiline rows={4} />
                <TextField
                  label="What are your ESG Interests?"
                  multiline
                  rows={4}
                />
                <TextField
                  label="Do you hold any board positions, if so which boards?"
                  multiline
                  rows={4}
                />
                <TextField
                  label="What are you passionate about?"
                  multiline
                  rows={4}
                />
              </Stack>
            </Stack>

          </PivotItem>
          {/* SUBSCRIPTIONS TAB */}
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
    </div>
  );
};

export default ClientProfileInfo;
