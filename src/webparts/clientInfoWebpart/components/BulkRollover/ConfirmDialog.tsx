import React from 'react'
import {
  ListView,
  IViewField,
  SelectionMode,
  GroupOrder,
  IGrouping,
} from "@pnp/spfx-controls-react/lib/ListView";
import styles from "../ClientInfoWebpart.module.scss";
import { MatterAndRolloverData } from './rolloverLogic'


const ConfirmDialog = (items: MatterAndRolloverData[], ) => {
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
    <div className={styles.listViewPortsForRollover}>
      <ListView
        items={items}
        viewFields={viewFields}
        // iconFieldName="FileRef"
        compact={true}
        selectionMode={SelectionMode.none}
        // selection={(selectionItem) => setPortalSelected(selectionItem)}
        // defaultSelection={defaultSelectedFromScreen2}
        showFilter={false}
        key="engagementPortals"
      />
    </div>
  );
};

export default ConfirmDialog
