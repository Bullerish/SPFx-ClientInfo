import * as React from "react";
import { useState, useEffect } from "react";
import { Modal, IDragOptions } from "office-ui-fabric-react/lib/Modal";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import { GlobalValues } from "../../Dataprovider/GlobalValue";
import { sp } from "@pnp/sp";

// parent container Manage Alerts component
const ManageAlerts = ({ spContext, isAlertModalOpen, onAlertModalHide }) => {
  const [subWebInfo, setSubWebInfo] = useState<any[]>([]);

  const hostUrl = window.location.host;
  console.log("hosturl: ", hostUrl);

  useEffect(() => {
    // get sub-portal information
    async function getSubwebs() {
      const subWebs = await sp.web
        .getSubwebsFilteredForCurrentUser()
        .select("Title", "ServerRelativeUrl", "Id")
        .orderBy("Created", false)();
      // console.table(subWebs);
      setSubWebInfo(subWebs);
    }

    getSubwebs();
  }, []);

  // will run only if subWebInfo is changed
  useEffect(() => {
    if (subWebInfo.length > 0) {
      // get current alerts set for user
      // TODO: grab ServerRelativeUrl from getSubwebs(), build below fetch with hostUrl var and ServerRelativeUrl to check if current user has an alert set on sub-portal (additional work to be done to check which list in sub-portal)
      fetch(
        `https://${hostUrl}${subWebInfo[0].ServerRelativeUrl}/_api/web/alerts`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      )
        .then((data) => {
          return data.json();
        })
        .then((alert) => {
          console.log(alert);
        });
    }
  }, [subWebInfo]);

  // using to test state updates
  useEffect(() => {
    console.log(subWebInfo);
  }, [subWebInfo]);

  return (
    <div>
      <Modal
        isOpen={isAlertModalOpen}
        onDismiss={() => onAlertModalHide(true)}
        isBlocking={false}
        // containerClassName={styles.container}
        // dragOptions={this.state.isDraggable ? this._dragOptions : undefined}
      >
        <DefaultButton onClick={() => onAlertModalHide(true)} text="Close" />
        <div>
          <p>
            Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas
            lorem nulla, malesuada ut sagittis sit amet, vulputate in leo.
            Maecenas vulputate congue sapien eu tincidunt. Etiam eu sem turpis.
            Fusce tempor sagittis nunc, ut interdum ipsum vestibulum non. Proin
            dolor elit, aliquam eget tincidunt non, vestibulum ut turpis. In hac
            habitasse platea dictumst. In a odio eget enim porttitor maximus.
            Aliquam nulla nibh, ullamcorper aliquam placerat eu, viverra et dui.
            Phasellus ex lectus, maximus in mollis ac, luctus vel eros. Vivamus
            ultrices, turpis sed malesuada gravida, eros ipsum venenatis elit,
            et volutpat eros dui et ante. Quisque ultricies mi nec leo ultricies
            mollis. Vivamus egestas volutpat lacinia. Quisque pharetra eleifend
            efficitur.
          </p>
          <p>
            Nam id mi justo. Nam vehicula vulputate augue, ac pretium enim
            rutrum ultricies. Sed aliquet accumsan varius. Quisque ac auctor
            ligula. Fusce fringilla, odio et dignissim iaculis, est lacus
            ultrices risus, vitae condimentum enim urna eu nunc. In risus est,
            mattis non suscipit at, mattis ut ante. Maecenas consectetur urna
            vel erat maximus, non molestie massa consequat. Duis a feugiat nibh.
            Sed a hendrerit diam, a mattis est. In augue dolor, faucibus vel
            metus at, convallis rhoncus dui.
          </p>
        </div>
      </Modal>
    </div>
  );
};

export default ManageAlerts;
