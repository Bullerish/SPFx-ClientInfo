import * as React from "react";
import {
  Toggle, Dropdown, Slider, TextField, Label, IDropdownOption
} from 'office-ui-fabric-react';
import { IAssuranceEngSplit } from "./ICreateEngagement";
import styles from "../ClientInfoWebpart.module.scss";
import { initializeIcons } from 'office-ui-fabric-react';

initializeIcons();
const SplitFrequencyOptions = [
  { key: 'Quarterly', text: 'Quarterly' },
  { key: 'Semi-Annual', text: 'Semi-Annual' },
  { key: 'Monthly', text: 'Monthly' },
  { key: 'Miscellaneous', text: 'Miscellaneous' },
];

class SplitValues {
  public Id;
  public Type;
  public Errormessage;
  public Value;
}

class AssuranceEngSplit extends React.Component<IAssuranceEngSplit> {

  public state = {
    disabled: true,
    maxval: 0,
    setSliderValue: 0,
    SelectedCategory: "",
    txtValues: [],
    fieldsArray: [],
    sliderVal: 0,
    isSplitNameExist: false,
    minVal: 0,
    splitToggleValue: false

  };


  public _onChangeSplitCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ sliderVal: 1, fieldsArray: [], txtValues: [], SelectedCategory: item.key }, () => {

      this.sliderOnChange(this.state.sliderVal);
      if (item.text == "Quarterly") {
        this.setState({ maxval: 4 });
      }
      else if (item.text == "Semi-Annual") {
        this.setState({ maxval: 2 });
      }
      else if (item.text == "Monthly") {
        this.setState({ maxval: 12 });
      }
      else if (item.text == "Miscellaneous") {
        this.setState({ maxval: 25 });
      }
      else {
        this.setState({ maxval: 0 });

      }
      this.props.OnSplitChange(this.state);
    });
  }

  public sliderOnChange = (value: number) => {
    var fieldsArray = [];
    if (this.state.SelectedCategory && this.state.SelectedCategory != "") {
      let prefix = "";
      let txtvalues = [];
      this.setState({ setSliderValue: value, sliderVal: value, fieldsArray: [] }, () => {
        if (this.state.SelectedCategory == "Quarterly") {
          prefix = "Q";
        }
        else if (this.state.SelectedCategory == "Semi-Annual") {
          prefix = "SEMI";
        }
        else if (this.state.SelectedCategory == "Monthly") {
          prefix = "M";
        }
        else if (this.state.SelectedCategory == "Miscellaneous") {
          prefix = "MISC";
        }

        this.setState({ txtValues: [] });
        for (var i = 0; i < value; i++) {
          let value1 = prefix + (i + 1).toString();
          let taxVal = new SplitValues();
          taxVal.Value = value1;
          taxVal.Id = "txt" + (i + 1).toString();
          txtvalues.push(taxVal);
          this.setState({ txtValues: txtvalues });
        }
        this.setState({ fieldsArray: fieldsArray });
      });
      this.props.OnSplitChange(this.state);
    }
  }

  public onChangeSecondTextFieldValue =
    (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
      if (!newValue || newValue.length <= 10) {
      }
    }

  public onBlurSplitName = (e, index) => {
    const regex = /^[0-9a-zA-Z]+$/;
    const { value, id } = e.target;
    let txtValues = [...this.state.txtValues];
    let item = { ...txtValues[index] };
    var isExist = txtValues.filter(t => t.Value.toString().toLowerCase() == value.toString().toLowerCase() && t.Id != id).length > 0;

    if (!value.match(regex) || value === "") {
      item['Value'] = "";
      item['Errormessage'] = "Only Alphanumeric values are allowed";
    }
    else if (isExist) {
      item['Value'] = "";
      item['Errormessage'] = "Already Exists";
    }
    else {
      item['Value'] = value;
      item['Errormessage'] = "";
    }

    txtValues[index] = item;

    this.setState({ txtValues }, () => {
      this.props.OnSplitChange(this.state);
    });
  }

  public onToggleChange = (ev, checked) => {
    if (checked) {
      this.setState({ splitToggleValue: true, disabled: false, sliderVal: 1, minVal: 1 }, () => {
        this.sliderOnChange(this.state.sliderVal);
        this.props.OnSplitChange(this.state);
      });

    } else {
      this.setState({ splitToggleValue: false, disabled: true, txtValues: [], sliderVal: 0, minVal: 0 }, () => {
        this.props.OnSplitChange(this.state);
      });
    }
  }

  public componentDidMount() {
    let { AsuranceSplitData } = this.props;
    this.setState(AsuranceSplitData);

  }

  public render(): React.ReactElement<IAssuranceEngSplit> {

    return (
      <div className={styles.splitOptions}>
        <div className={styles.splitColumns}>

          <Toggle
            label="Split Options"
            onText="Split Engagement"
            offText="Split Engagement"
            onChange={(ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
              this.onToggleChange(ev, checked);
            }}
            checked={this.state.splitToggleValue}
          >
          </Toggle>
          <Dropdown
            placeholder="Select split frequency"
            label="Split Category"
            options={SplitFrequencyOptions}
            disabled={this.state.disabled}
            onChange={this._onChangeSplitCategory}
            selectedKey={this.state.SelectedCategory}
          />
          <Slider label="Quantity" value={this.state.sliderVal} min={this.state.minVal} max={this.state.maxval} onChanged={(event: any, value: number, range?: [number, number]) => this.sliderOnChange(value)} showValue disabled={this.state.disabled} />
        </div>

        <div className={styles.quarterColumns}>
          <Label>Split Name</Label>
          {this.state.txtValues.map(
            (ctrl, index) => <TextField maxLength={10} onBlur={(e) => this.onBlurSplitName(e, index)} id={ctrl.Id} errorMessage={ctrl.Errormessage} value={ctrl.Value} />
          )}
        </div>

      </div>
    );

  }


}
export default AssuranceEngSplit;
