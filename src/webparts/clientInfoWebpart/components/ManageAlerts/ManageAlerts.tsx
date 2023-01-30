import * as React from "react";
import { useState } from "react";
import { Modal, IDragOptions } from "office-ui-fabric-react/lib/Modal";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import { GlobalValues } from "../../Dataprovider/GlobalValue";





// parent container Manage Alerts component
const ManageAlerts = ({spContext, isAlertModalOpen, onAlertModalHide}) => {
  const [isModalOpen, setIsModalOpen] = useState<boolean>(isAlertModalOpen);


  // open the modal window
  function showModal(): void {
    setIsModalOpen(true);
  }

  // close the modal window
  function closeModal(): void {
    setIsModalOpen(false);
  }



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
            efficitur.{" "}
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
