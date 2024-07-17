import React from 'react';
import {
  ListView,
  IViewField,
  SelectionMode,
} from "@pnp/spfx-controls-react/lib/ListView";
import styles from "../ClientInfoWebpart.module.scss";
import { MatterAndCreationData } from './creationLogic';

interface ConfirmDialogProps {
  items: MatterAndCreationData[];
}

const ConfirmDialog: React.FC<ConfirmDialogProps> = ({ items }) => {
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
    {
      name: "siteOwner",
      displayName: "Site Owner",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
    {
      name: "newMatterPortalExpirationDate",
      displayName: "Expiration Date",
      sorting: false,
      minWidth: 100,
      maxWidth: 225,
      isResizable: true,
    },
  ];

  return (
    <div className={styles.listViewPortsForCreation}>
      <ListView
        items={items}
        viewFields={viewFields}
        compact={true}
        selectionMode={SelectionMode.none}
        showFilter={false}
        key="engagementPortals"
      />
    </div>
  );
};

export default ConfirmDialog;
