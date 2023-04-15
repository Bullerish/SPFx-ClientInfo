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
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  CheckboxVisibility,
  SelectionMode,
  IObjectWithKey,
} from "office-ui-fabric-react/lib/DetailsList";
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
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';

interface IDropdownControlledState {
  jobFunctionItem?: { key: string | number | undefined };
  jobRoleItem?: { key: string | number | undefined };
}

interface IServices {
  text: string;
  checked: boolean;
}

const jobFunctionOptions = [
  { key: 'Accounting', text: 'Accounting' },
  { key: 'Administration', text: 'Administration' },
  { key: 'Analytics', text: 'Analytics' },
  { key: 'Compliance', text: 'Compliance' },
  { key: 'Executive', text: 'Executive' },
  { key: 'Finance', text: 'Finance' },
  { key: 'Human Resources', text: 'Human Resources' },
  { key: 'IT', text: 'IT' },
  { key: 'Legal', text: 'Legal' },
  { key: 'Operations', text: 'Operations' },
  { key: 'Risk', text: 'Risk' },
  { key: 'Sales/Marketing', text: 'Sales/Marketing' },
  { key: 'Tax', text: 'Tax' },
  { key: 'Other', text: 'Other' },
];

const jobRoleOptions = [
  { key: 'Board Member', text: 'Board Member' },
  { key: 'Csuite', text: 'Csuite' },
  { key: 'VP / Director', text: 'VP / Director' },
  { key: 'Manager', text: 'Manager' },
  { key: 'Staff', text: 'Staff' },
  { key: 'Student', text: 'Student' },
  { key: 'Other', text: 'Other' },
];

const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } }
};

const profileContainerProps: Partial<IStackProps> = {
  styles: { root: { width: 700, marginTop: 20 } }
};

const textAreasProps: Partial<IStackProps> = {
  styles: { root: { width: 650 } }
};



const ClientProfileInfo = ({
  spContext,
  isClientProfileInfoModalOpen,
  onClientProfileInfoModalHide,
}): React.ReactElement => {
  const [jobFunctionItem, setJobFunctionItem] = useState({ key: undefined, text: '' });

  // for checkboxes
  const [services, setServices] = useState<IServices[]>([
    {text: 'Accounting/Assurance', checked: false},
    {text: 'Advisory', checked: false},
    {text: 'Private Client Services', checked: false},
    {text: 'Solutions for C-Suite Executives', checked: false},
    {text: 'Tax', checked: false},
  ]);

  useEffect(() => {
    console.log('triggering useEffect services');
    console.log(services);
  }, [services]);


  // set services checkboxes and state
  const onServicesChange = (text: string, value) => {

    let servicesTempArr = services;

    servicesTempArr.forEach((e) => {
      if (text === e.text && value) {
        console.log('inside foreach logging val:: ', e);
        e.checked = true;
      }

      if (text === e.text && !value) {
        e.checked = false;
      }
    });

    setServices([...servicesTempArr]);
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
          className: styles.manageAlerts,
        }}
        // styles={{ root: { maxHeight: 700 } }}
      >
        <Pivot>
          <PivotItem
            headerText="Profile"
            headerButtonProps={{
              "data-order": 1,
              "data-title": "Profile Title",
            }}
          >

            {/* <div style={{ display: 'flex', justifyContent: 'center', width: '100%' }}> */}
            <Stack gap={20} horizontalAlign="center" {...profileContainerProps}>
              <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 50 }} styles={{ root: { width: 650 } }}>


                <Stack {...columnProps}>

                  <TextField label="Full Name" />

                  <Dropdown
                    label="Job Role"
                    selectedKey={jobFunctionItem ? jobFunctionItem.key : undefined}
                    // onChange={this._onChange}
                    placeholder="Select an option"
                    options={jobRoleOptions}
                    styles={{ dropdown: { width: 300 } }}
                  />

                </Stack>

                <Stack {...columnProps} >

                  <TextField label="Title" />

                  <Dropdown
                    label="Job Function"
                    selectedKey={jobFunctionItem ? jobFunctionItem.key : undefined}
                    // onChange={this._onChange}
                    placeholder="Select an option"
                    options={jobFunctionOptions}
                    styles={{ dropdown: { width: 300 } }}
                  />

                </Stack>



              </Stack>

              <Stack gap={20} {...textAreasProps}>

                <TextField label="Mailing Address" multiline rows={4} />
                <TextField label="What are your ESG Interests?" multiline rows={4} />
                <TextField label="Do you hold any board positions, if so which boards?" multiline rows={4} />
                <TextField label="What are you passionate about?" multiline rows={4} />

              </Stack>
            </Stack>
            {/* </div> */}

          </PivotItem>
          <PivotItem headerText="Subscriptions">

            <div>SERVICES</div>

            {services.map(e => (
              <Checkbox label={e.text} checked={e.checked} onChange={(ev, value) => onServicesChange(e.text, value)} />
            ))}

          </PivotItem>
          <PivotItem headerText="Contacts">
            <div>Contacts</div>
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
