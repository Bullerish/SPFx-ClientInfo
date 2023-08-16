
import * as XLSX from 'xlsx';

const headerColumns = ['K1 Partner Folder Name', 'Email Address (example: email1@domain.com;email2@domain.com)'];

export class K1ImportCheck {

    public static validateK1File(file) {
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        //TODO: ANALYZE THIS FILE        
        const reader = new FileReader();
        //reader.readAsText(file);          
        //reader.onload = (event: any) => {
        // var data = event.target.result;
        let data = XLSX.read(file, { type: 'binary', raw: true });
        const wsname = data.SheetNames[0];
        const ws = data.Sheets[wsname];
        /* Convert array to json*/
        const dataParse = XLSX.utils.sheet_to_json(ws, { header: 1 });
        let isValidColumns = false;
        var errorMessages = { ColumnLengthError: "", ColumnError: "", ColumnRequiredError: [], ColumnTypeLengthError: [] };
        //const TypeLengthValidationRule = this.state.TypeLengthValidationRule;
        if (dataParse && dataParse.length > 0) {
            let headerColumn: any = dataParse[0];
            if (headerColumn.length == headerColumns.length) {
                /**Validation for No. of column and Column names */
                for (let i = 0; i < headerColumn.length; i++) {
                    if (headerColumn[i] == headerColumns[i]) {
                        isValidColumns = true;
                    } else {
                        isValidColumns = false;
                        errorMessages["ColumnError"] = `column ${headerColumns[i]} is not found`;
                        break;
                    }
                }
                if (isValidColumns) {
                    let folderCount = 0;
                    /**Validation for required column values */
                    for (let i = 1; i < dataParse.length; i++) {
                        let item = dataParse[i];
                        /*if (!this.isValidRow(item)) {
                            console.log(`${i + 1} is blank`);
                            continue;
                        }*/
                        folderCount++;
                        let colNames = [];
                        headerColumns.forEach(checkIndex => {
                            if (!item[checkIndex] || (item[checkIndex] && item[checkIndex] == "")) {
                                colNames.push(headerColumns[checkIndex]);
                            }
                        });
                        if (colNames.length > 0) {
                            errorMessages["ColumnRequiredError"].push(`Row ${i + 1}: Missing value for ${colNames.join(', ')}`);
                        }
                        console.log('item is', item);
                        /*
                        let _self = this;
                        Object.keys(TypeLengthValidationRule).forEach(async (key) => {
                            let rules = TypeLengthValidationRule[key];
                            let itemValue = item[rules.index];

                            if (itemValue !== null && itemValue !== undefined && itemValue !== "") {
                                if ("maxLength" in rules) {
                                    if (itemValue.length > rules.maxLength) {
                                        errorMessages["ColumnTypeLengthError"].push(`Row ${i + 1}: ${key} exceeds maximum characters (${rules.maxLength})`);
                                    }
                                }
                                if ("isAlphanumeric" in rules) {
                                    if (!_self.CheckAlphaNumeric(itemValue)) {
                                        errorMessages["ColumnTypeLengthError"].push(`Row ${i + 1}: ${key} should be alphanumeric`);
                                    }
                                }
                                if ("allowOnly" in rules) {
                                    if (rules.allowOnly.indexOf(itemValue) === -1) {
                                        errorMessages["ColumnTypeLengthError"].push(`Row ${i + 1}: ${key} has an invalid value`);
                                    }
                                }
                                if ('isDate' in rules) {
                                    if (!_self.CheckValidDate(itemValue)) {
                                        errorMessages["ColumnTypeLengthError"].push(`Row ${i + 1}: ${key} doesn't have a valid date format (MM/DD/YYYY)`);
                                    }
                                }
                                if ('isEmail' in rules) {
                                    if (!_self.CheckValidEmail(itemValue)) {
                                        errorMessages["ColumnTypeLengthError"].push(`Row ${i + 1}: ${key} has an invalid email value`);
                                    }
                                }
                            }
                        });*/
                    }
                }
            }
            else {
                errorMessages["ColumnLengthError"] = "number of columns are not valid";
            }
        }
        console.log("errorMessages", errorMessages);
        if (errorMessages) {
            return false;
        }
        else {
            return true;
        }
        //   this.setState({ ErrorMessages: errorMessages, IsValidColumns: isValidColumns, TotalRows: dataParse, isFileUpload: true });
    }


    // Used to check K1 CSV rows
    private isValidRow(row) {
        if (!row || row.length == 0)
            return false;
        var inValidRowCount = 0;
        row.forEach(column => {
            if (!column || column.trim() == "")
                inValidRowCount++;
        });
        return inValidRowCount == 0 ? true : false;
    }

    private CheckValidEmail(value) {
        let exp = /^([a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})*$/;
        if (value && value.match(exp)) {
            return true;
        } else {
            return false;
        }
    }

}
