import React, { useState } from 'react';
import { DefaultButton, PrimaryButton, Dialog, DialogType, DialogFooter, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import { getMatterNumbersForClientSite, createEngagementSubportals } from './creationLogic';

const BulkCreation: React.FC = () => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [portalType, setPortalType] = useState<string>('');
  const [team, setTeam] = useState<string>('');
  const [selectedEngagements, setSelectedEngagements] = useState<any[]>([]);
  const [step, setStep] = useState<number>(1);

  const portalTypeOptions: IDropdownOption[] = [
    { key: 'workflow', text: 'Workflow' },
    { key: 'fileExchange', text: 'File Exchange' },
  ];

  const teamOptions: IDropdownOption[] = [
    { key: 'assurance', text: 'Assurance' },
    { key: 'tax', text: 'Tax' },
    { key: 'advisory', text: 'Advisory' },
  ];

  const openModal = () => setIsModalOpen(true);
  const closeModal = () => setIsModalOpen(false);
  
  const handlePortalTypeChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
    setPortalType(item.key as string);
  };

  const handleTeamChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
    setTeam(item.key as string);
  };

  const handleEngagementSelection = (engagement: any) => {
    const index = selectedEngagements.findIndex((e) => e.ID === engagement.ID);
    if (index === -1) {
      setSelectedEngagements([...selectedEngagements, engagement]);
    } else {
      const newSelections = [...selectedEngagements];
      newSelections.splice(index, 1);
      setSelectedEngagements(newSelections);
    }
  };

  const handleNextStep = () => {
    setStep(step + 1);
  };

  const handlePrevStep = () => {
    setStep(step - 1);
  };

  const handleCreatePortals = async () => {
    await createEngagementSubportals(selectedEngagements, portalType, team);
    closeModal();
    alert('Thank you. Your portals are in the process of being created. You will receive an email confirmation shortly when your portals are active. Please close this window.');
  };

  return (
    <div>
      <DefaultButton text="Bulk subportal creation" onClick={openModal} />
      <Dialog
        hidden={!isModalOpen}
        onDismiss={closeModal}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Bulk Subportal Creation',
          subText: 'Select Portal Type, Team, and Engagements for bulk creation',
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        {step === 1 && (
          <div>
            <Dropdown
              label="Portal Type"
              options={portalTypeOptions}
              onChange={handlePortalTypeChange}
              selectedKey={portalType}
            />
            <Dropdown
              label="Team"
              options={teamOptions}
              onChange={handleTeamChange}
              selectedKey={team}
            />
            <div>
              {/* List of engagements with multi-select functionality */}
              {/* Example engagement list rendering */}
              {mockEngagements.map((engagement) => (
                <div
                  key={engagement.ID}
                  onClick={() => handleEngagementSelection(engagement)}
                  style={{ background: selectedEngagements.includes(engagement) ? '#eaeaea' : 'white' }}
                >
                  {engagement.Title}
                </div>
              ))}
            </div>
            <DialogFooter>
              <PrimaryButton text="Next" onClick={handleNextStep} />
              <DefaultButton text="Cancel" onClick={closeModal} />
            </DialogFooter>
          </div>
        )}
        {step === 2 && (
          <div>
            <p>Review your selections:</p>
            {/* Display selected engagements and other details */}
            {selectedEngagements.map((engagement) => (
              <div key={engagement.ID}>{engagement.Title}</div>
            ))}
            <DialogFooter>
              <PrimaryButton text="Create Portals" onClick={handleCreatePortals} />
              <DefaultButton text="Back" onClick={handlePrevStep} />
            </DialogFooter>
          </div>
        )}
      </Dialog>
    </div>
  );
};

// Example data, replace with actual data fetching logic
const mockEngagements = [
  { ID: '1', Title: 'Engagement 1' },
  { ID: '2', Title: 'Engagement 2' },
  { ID: '3', Title: 'Engagement 3' },
];

export default BulkCreation;
