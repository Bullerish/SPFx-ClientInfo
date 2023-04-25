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
import { IItemAddResult } from "@pnp/sp/items";

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

const ClientProfileInfo = ({spContext, isClientProfileInfoModalOpen, onClientProfileInfoModalHide, showHideAlertIcon, showHideToast, isDontRemind}): React.ReactElement => {
  // state for currentUser and Webs instances
  const [currentUser, setCurrentUser] = useState<ISiteUserInfo>(null);
  const [userItemData, setUserItemData] = useState<any[]>([]);
  const [itemID, setItemID] = useState<number>(null);

  // states for confirmation and status dialogs
  const [isSubmissionSuccessful, setIsSubmissionSuccessful] = useState<boolean>(null);
  const [confirmDialogHidden, setConfirmDialogHidden] = useState<boolean>(true);
  const [statusDialogHidden, setStatusDialogHidden] = useState<boolean>(true);

  // flags for "!" icon and reminder toast
  const [isComplete, setIsComplete] = useState<boolean>(null);
  const [reminder, setReminder] = useState<boolean>(false);

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
    { text: "Government - Audit/Accounting", checked: false },
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

  // object payload to be submitted in add/update pnpjs calls
  let payload = {
    Title: title,
    FullName: fullName,
    JobLevel: jobRoleItem.key,
    JobFunction: jobFunctionItem.key,
    MailingAddress: mailingAddress,
    BoardPositions: boardPositions,
    Passions: passions,
    Services: {}, // need to update to proper format prior to submitting data otherwise 400 will occur
    Sectors: {},
    OtherInterests: {},
    ESGInterests: esgInterests,
    Contacts: contacts,
    UserLoginName: currentUser !== null ? currentUser.LoginName : '',
    isComplete: isComplete,
    Reminder: reminder,
  };

  // this useEffect will run only once after initially page load/component mount
  useEffect(() => {
    const siteWebVal = Web(GlobalValues.SiteURL);

    const getCurrentUser = async () => {
      const userInfo = await siteWebVal.currentUser();
      setCurrentUser(userInfo);
    };

    getCurrentUser();

  }, []);

  // check list for current logged in user list item. Set the usetItemData state variable with results
  useEffect(() => {
    console.log(currentUser);

    // check to see if user has existing item, if so bring it back
    const checkSetUserItem = async () => {
      const hubWeb = Web(GlobalValues.HubSiteURL);

      if (currentUser !== null) {
        const loginName = currentUser.LoginName;
        const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.select('ID', 'Title', 'FullName', 'JobLevel', 'JobFunction', 'MailingAddress', 'BoardPositions', 'Passions', 'Services', 'Sectors', 'OtherInterests', 'ESGInterests', 'Contacts', 'UserLoginName', 'isComplete', 'Reminder').filter(`UserLoginName eq '${loginName}'`).get();
        console.log('logging userItem:: ', userItem);
        setUserItemData(userItem);
      }

    };

    if (currentUser !== null) {
      checkSetUserItem();
    }

  }, [currentUser]);

  // runs when userItemData has been populated with user's list item data
  useEffect(() => {
    console.log('in userItemData useEffect::');
    // console.log('logging userItemData fullName:: ', userItemData[0].FullName);
    let newServices = services;
    let newSectors = sectors;
    let newOtherInterests = otherInterests;
    // TODO: will have to update form complete/reminder criteria based on marketing/courtney feedback
    if (!userItemData.length) {

      setIsComplete(false);
      setReminder(true);
      showHideAlertIcon(true);
      showHideToast(true);

    } else {
      // set all states here to populate form fields
      setFullName(userItemData[0].FullName);
      setTitle(userItemData[0].Title);
      setJobRoleItem({ key: userItemData[0].JobLevel, text: userItemData[0].JobLevel });
      setJobFunctionItem({ key: userItemData[0].JobFunction, text: userItemData[0].JobFunction });
      setMailingAddress(userItemData[0].MailingAddress);
      setEsgInterests(userItemData[0].ESGInterests);
      setBoardPositions(userItemData[0].BoardPositions);
      setPassions(userItemData[0].Passions);
      setContacts(userItemData[0].Contacts);

      // additional logic needed to set checkboxes
      if (userItemData[0].Services) {
        newServices.forEach(item => {if (userItemData[0].Services.indexOf(item.text) !== -1) {item.checked = true;}});
      }
      if (userItemData[0].Sectors) {
        newSectors.forEach(item => {if (userItemData[0].Sectors.indexOf(item.text) !== -1) {item.checked = true;}});
      }
      if (userItemData[0].OtherInterests) {
        newOtherInterests.forEach(item => {if (userItemData[0].OtherInterests.indexOf(item.text) !== -1) {item.checked = true;}});
      }

      setServices(newServices);
      setSectors(newSectors);
      setOtherInterests(newOtherInterests);

      setIsComplete(userItemData[0].isComplete);
      setReminder(userItemData[0].Reminder);
      setItemID(userItemData[0].ID);
    }
  }, [userItemData]);

  useEffect(() => {
    console.log('in itemID useEffect::');
    console.log('logging reminder state:: ', reminder);
    // console.log('logging contacts state in itemID useEffect:: ', contacts);
    // factor logic here to determine whether or not we display the alert icon on the main client info component
    if (itemID !== null) {
      if (
        (fullName !== '' && fullName !== null) &&
        (title !== '' && title !== null) &&
        jobRoleItem.key &&
        jobFunctionItem.key &&
        (mailingAddress !== '' && mailingAddress !== null) &&
        (esgInterests !== '' && esgInterests !== null) &&
        (boardPositions !== '' && boardPositions !== null) &&
        (passions !== '' && passions !== null) &&
        (contacts !== '' && contacts !== null) &&
        services.length &&
        sectors.length &&
        otherInterests.length
      ) {
        console.log('in passing block of itemID useEffect::');
        setIsComplete(true);
        showHideAlertIcon(false);
        showHideToast(false);
      } else {
        console.log('in failing block of itemID useEffect::');
        setIsComplete(false);
        showHideAlertIcon(true);
        if (reminder) {
          console.log('in passing block of reminder check::');
          showHideToast(true);
        } else {
          console.log('in failing block of reminder check::');
          showHideToast(false);
        }
      }
    }
  }, [itemID]);

  useEffect(() => {
    const hubWeb = Web(GlobalValues.HubSiteURL);
    const loginName = currentUser !== null ? currentUser.LoginName : null;
    console.log('running isDontRemind useEffect:: ', isDontRemind);

    const runUpdateOnListItem = async () => {

      if (isDontRemind && itemID === null) {
        const addItemResult: IItemAddResult = await hubWeb.lists.getByTitle(clientProfileListName).items.add({
          Reminder: false,
          UserLoginName: loginName
        });
        // TODO: check for item submission successful, then run getitembyid to fetch the newly updated list item, set the result in state var userItemData
        if (currentUser !== null && addItemResult !== null) {

          const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.select('ID', 'Title', 'FullName', 'JobLevel', 'JobFunction', 'MailingAddress', 'BoardPositions', 'Passions', 'Services', 'Sectors', 'OtherInterests', 'ESGInterests', 'Contacts', 'UserLoginName', 'isComplete', 'Reminder').filter(`UserLoginName eq '${loginName}'`).get();
          console.log('logging userItem:: ', userItem);
          setUserItemData(userItem);
        }
      } else if (isDontRemind && itemID !== null) {
        const updatedItemResult = await hubWeb.lists.getByTitle(clientProfileListName).items.getById(itemID).update({
          Reminder: false,
          UserLoginName: loginName
        });
        // TODO: check for item submission successful, then run getitembyid to fetch the newly updated list item, set the result in state var userItemData
        if (currentUser !== null && updatedItemResult !== null) {

          const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.select('ID', 'Title', 'FullName', 'JobLevel', 'JobFunction', 'MailingAddress', 'BoardPositions', 'Passions', 'Services', 'Sectors', 'OtherInterests', 'ESGInterests', 'Contacts', 'UserLoginName', 'isComplete', 'Reminder').filter(`UserLoginName eq '${loginName}'`).get();
          console.log('logging userItem:: ', userItem);
          setUserItemData(userItem);
        }

      }
    };

    runUpdateOnListItem();

  }, [isDontRemind]);

  // set services checkboxes and state
  const onServicesChange = (text: string, value) => {
    let servicesTempArr = services;

    servicesTempArr.forEach((e) => {
      if (text === e.text && value) {
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
    setConfirmDialogHidden(true);
    onClientProfileInfoModalHide();
  };

  // resetting
  const resetState = () => {
    let newServices = services;
    let newSectors = sectors;
    let newOtherInterests = otherInterests;

    if (userItemData.length) {
      // run for update & set existing user item as state
      setFullName(userItemData[0].FullName);
      setTitle(userItemData[0].Title);
      setJobRoleItem({ key: userItemData[0].JobLevel, text: userItemData[0].JobLevel });
      setJobFunctionItem({ key: userItemData[0].JobFunction, text: userItemData[0].JobFunction });
      setMailingAddress(userItemData[0].MailingAddress);
      setEsgInterests(userItemData[0].ESGInterests);
      setBoardPositions(userItemData[0].BoardPositions);
      setPassions(userItemData[0].Passions);
      setContacts(userItemData[0].Contacts);

      // additional logic needed to set checkboxes
      if (userItemData[0].Services) {
        newServices.forEach(item => {if (userItemData[0].Services.indexOf(item.text) !== -1) {item.checked = true;}});
      }
      if (userItemData[0].Sectors) {
        newSectors.forEach(item => {if (userItemData[0].Sectors.indexOf(item.text) !== -1) {item.checked = true;}});
      }
      if (userItemData[0].OtherInterests) {
        newOtherInterests.forEach(item => {if (userItemData[0].OtherInterests.indexOf(item.text) !== -1) {item.checked = true;}});
      }

      console.log('logging Services from userItemData:: ', userItemData[0].Services);

      if (userItemData[0].Services === null) {
        setServices([
          { text: "Accounting/Assurance", checked: false },
          { text: "Advisory", checked: false },
          { text: "Private Client Services", checked: false },
          { text: "Solutions for C-Suite Executives", checked: false },
          { text: "Tax", checked: false },
        ]);
      } else {
        setServices(newServices);
      }

      if (userItemData[0].Sectors === null) {
        setSectors([
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
          { text: "Government - Audit/Accounting", checked: false },
          { text: "Government - Compliance and Monitoring", checked: false },
          { text: "Government - Emergency Management", checked: false },
        ]);
      } else {
        setSectors(newSectors);
      }

      if (userItemData[0].OtherInterests === null) {
        setOtherInterests([
          { text: "Alumni Events", checked: false },
          { text: "CPE Offerings", checked: false },
          { text: "Executive Women's Forum Events", checked: false },
        ]);
      } else {
        setOtherInterests(newOtherInterests);
      }

      setIsComplete(userItemData[0].isComplete);
      setReminder(userItemData[0].Reminder);
      setItemID(userItemData[0].ID);

    } else {
      // run for new
      setFullName('');
      setTitle('');
      setJobRoleItem({
        key: undefined,
        text: "",
      });
      setJobFunctionItem({
        key: undefined,
        text: "",
      });
      setMailingAddress('');
      setEsgInterests('');
      setBoardPositions('');
      setPassions('');
      setContacts('');

      setServices([
        { text: "Accounting/Assurance", checked: false },
        { text: "Advisory", checked: false },
        { text: "Private Client Services", checked: false },
        { text: "Solutions for C-Suite Executives", checked: false },
        { text: "Tax", checked: false },
      ]);
      setSectors([
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
        { text: "Government - Audit/Accounting", checked: false },
        { text: "Government - Compliance and Monitoring", checked: false },
        { text: "Government - Emergency Management", checked: false },
      ]);
      setOtherInterests([
        { text: "Alumni Events", checked: false },
        { text: "CPE Offerings", checked: false },
        { text: "Executive Women's Forum Events", checked: false },
      ]);

    }

    onClientProfileInfoModalHide();

  };


  // TODO: need to update for add/update
  const submitUserProfileInfo = async () => {
    const hubWeb = Web(GlobalValues.HubSiteURL);
    // set empty temp arrs to populate
    let servicesStringArr: string[] = [];
    let sectorsStringArr: string[] = [];
    let otherInterestsStringArr: string[] = [];

    // filter on state arrs for only "checked" items, push the text of each object to the respective temp arr
    services.filter(e => e.checked === true).forEach(e => servicesStringArr.push(e.text));
    sectors.filter(e => e.checked === true).forEach(e => sectorsStringArr.push(e.text));
    otherInterests.filter(e => e.checked === true).forEach(e => otherInterestsStringArr.push(e.text));

    console.log('logging services arr to send in payload:: ', servicesStringArr);

    // update payload props with new temp arrs from above
    payload.Services = {results: servicesStringArr};
    payload.Sectors = {results: sectorsStringArr};
    payload.OtherInterests = {results: otherInterestsStringArr};

    // determine to set payload vars, alert icon, and reminder notification
    if (
      (fullName !== '' && fullName !== null) &&
      (title !== '' && title !== null) &&
      jobRoleItem.key &&
      jobFunctionItem.key &&
      (mailingAddress !== '' && mailingAddress !== null) &&
      (esgInterests !== '' && esgInterests !== null) &&
      (boardPositions !== '' && boardPositions !== null) &&
      (passions !== '' && passions !== null) &&
      (contacts !== '' && contacts !== null) &&
      services.length &&
      sectors.length &&
      otherInterests.length
    ) {
      console.log('in passing block of itemID useEffect::');
      setIsComplete(true);
      payload.isComplete = true;
      payload.Reminder = false;
      showHideAlertIcon(false);
      showHideToast(false);
    } else {
      console.log('in failing block of itemID useEffect::');
      setIsComplete(false);
      payload.isComplete = false;
      // payload.Reminder = true;
      showHideAlertIcon(true);
      if (reminder) {
        console.log('in passing block of reminder check::');
        showHideToast(true);
        payload.Reminder = true;
      } else {
        console.log('in failing block of reminder check::');
        showHideToast(false);
        payload.Reminder = false;
      }
    }


    if (!userItemData.length) {
      // add new item to list
      console.log('logging payload:: ', payload);
      const addItemResult: IItemAddResult = await hubWeb.lists.getByTitle(clientProfileListName).items.add(payload);

      // set isSubmissionSuccessful flag to true or false based on if item was added successfully or not
      console.log('logging addItemResult:: ', addItemResult);
      if (addItemResult.data) {
        setIsSubmissionSuccessful(true);
        setStatusDialogHidden(false);



        const loginName = currentUser.LoginName;
        const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.select('ID', 'Title', 'FullName', 'JobLevel', 'JobFunction', 'MailingAddress', 'BoardPositions', 'Passions', 'Services', 'Sectors', 'OtherInterests', 'ESGInterests', 'Contacts', 'UserLoginName', 'isComplete', 'Reminder').filter(`UserLoginName eq '${loginName}'`).get();
        console.log('logging userItem:: ', userItem);
        setUserItemData(userItem);

      } else {
        setIsSubmissionSuccessful(false);
        setStatusDialogHidden(false);
      }
    } else {
      // update existing item in list
      console.log('logging payload:: ', payload);
      const updatedItemResult = await hubWeb.lists.getByTitle(clientProfileListName).items.getById(itemID).update(payload);

      console.log('logging updatedItemResult:: ', updatedItemResult);
      if (updatedItemResult.data) {
        setIsSubmissionSuccessful(true);
        setStatusDialogHidden(false);

        // // testing
        // if (
        //   (fullName !== '' && fullName !== null) &&
        //   (title !== '' && title !== null) &&
        //   jobRoleItem.key &&
        //   jobFunctionItem.key &&
        //   (mailingAddress !== '' && mailingAddress !== null) &&
        //   (esgInterests !== '' && esgInterests !== null) &&
        //   (boardPositions !== '' && boardPositions !== null) &&
        //   (passions !== '' && passions !== null) &&
        //   (contacts !== '' && contacts !== null) &&
        //   services.length &&
        //   sectors.length &&
        //   otherInterests.length
        // ) {
        //   console.log('in passing block of itemID useEffect::');
        //   setIsComplete(true);
        //   showHideAlertIcon(false);
        //   showHideToast(false);
        // } else {
        //   console.log('in failing block of itemID useEffect::');
        //   setIsComplete(false);
        //   showHideAlertIcon(true);
        //   if (reminder) {
        //     console.log('in passing block of reminder check::');
        //     showHideToast(true);
        //   } else {
        //     console.log('in failing block of reminder check::');
        //     showHideToast(false);
        //   }
        // }

        const loginName = currentUser.LoginName;
        const userItem = await hubWeb.lists.getByTitle(clientProfileListName).items.select('ID', 'Title', 'FullName', 'JobLevel', 'JobFunction', 'MailingAddress', 'BoardPositions', 'Passions', 'Services', 'Sectors', 'OtherInterests', 'ESGInterests', 'Contacts', 'UserLoginName', 'isComplete', 'Reminder').filter(`UserLoginName eq '${loginName}'`).get();
        console.log('logging userItem:: ', userItem);
        setUserItemData(userItem);
      } else {
        setIsSubmissionSuccessful(false);
        setStatusDialogHidden(false);
      }

    }

  };

  return (
    <div>
      <Dialog
        hidden={!isClientProfileInfoModalOpen}
        onDismiss={resetState}
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
                  <TextField label="Full Name" value={fullName} required onChange={(ev, newValue) => setFullName(newValue)} errorMessage={(fullName !== '' && fullName !== null) ? null : 'This field is required'} />

                  <Dropdown
                    label="Job Level"
                    selectedKey={
                      jobRoleItem ? jobRoleItem.key : undefined
                    }
                    onChange={(ev, item) => setJobRoleItem(item)}
                    placeholder="Select an option"
                    options={jobRoleOptions}
                    required
                    errorMessage={jobRoleItem.key ? null : 'This field is required'}
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
                    required
                    errorMessage={jobFunctionItem.key ? null : 'This field is required'}
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
                value={contacts}
                onChange={(ev, newValue) => setContacts(newValue)}
              />
            </div>

          </PivotItem>
        </Pivot>
        <DialogFooter>
          <Text style={{ color: 'red' }}>{((fullName !== '' && fullName !== null) && jobRoleItem.key && jobFunctionItem.key) ? null : '*Please fill out required fields to submit'}</Text>
          <PrimaryButton
            className={styles.primaryButton}
            onClick={() => setConfirmDialogHidden(false)}
            text="Save"
            disabled={((fullName !== '' && fullName !== null) && jobRoleItem.key && jobFunctionItem.key) ? false : true}
          />
          <DefaultButton
            className={styles.defaultButton}
            onClick={resetState}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
      {/* confirm dialog component. Modal/dialog window will open */}
      <ConfirmDialog confirmDialogHidden={confirmDialogHidden} onSetConfirmDialogHidden={onSetConfirmDialogHidden} onConfirmSubmission={submitUserProfileInfo} />
      {/* TODO: need to update logic after item submission to display Status Dialog component */}
      <StatusDialog isSubmissionSuccessful={isSubmissionSuccessful} statusDialogHidden={statusDialogHidden} onSetStatusDialogHidden={onSetStatusDialogHidden} />
    </div>
  );
};

export default ClientProfileInfo;
