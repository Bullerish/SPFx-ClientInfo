
import CSVFileValidator from 'csv-file-validator';

const config: any = {
    headers: [
        {
            name: 'K1 Partner Folder Name',
            required: true,
            requiredError: (headerName, rowNumber, columnNumber) => {
                return `${headerName} is missing in row ${rowNumber}.`;
            }
        },
        {
            name: 'Email Address (example: email1@domain.com;email2@domain.com)',
            required: true,
            validate: (emails) => {
                let testEmails = true;
                let emailsplit = emails.split(";"); 
                for (let i = 0; i < emailsplit.length; i++) {
                    let element = emailsplit[i];
                    element = element.replace(/\n/g, ''); // remove line breaks
                    element = element.trim(); // remove whitespace
                    if (element != "") {
                      //test for valid email
                      const emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
                      let emailTest = emailPattern.test(element);          
                      if (emailTest == false) {
                        testEmails = false;
                      }
                    }
                }
                return testEmails;
            },
            requiredError: (headerName, rowNumber, columnNumber) => {
                return `${headerName} is missing in row ${rowNumber}.`;
            },
            validateError: (headerName, rowNumber, columnNumber) => {
                return `${headerName} has invalid values in row ${rowNumber}.`;
            }
        }],
    isHeaderNameOptional: false,
    isColumnIndexAlphabetic: false
};

export class K1ImportCheck {    

    public static async validateK1File(file) {
        let errorMessage = [];
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        const filename = myfile.name;
        const extension = filename.substring(filename.lastIndexOf('.'));
        if (extension != ".csv") {
            errorMessage.push("Invalid file type.  Upload CSV files only.");
        }
        else {            
            // ANALYZE CSV FILE   
            let csvData = await CSVFileValidator(file, config);
            let csvContent= csvData.data; // Array of data          
            if (csvContent.length == 0) {
                errorMessage.push("This file is empty.  Please upload a valid csv file.");
            }
            else {
                let csvErrors = csvData.inValidData; // Array of error messages
                //console.log('csvInvalid', csvErrors);
                if (csvErrors.length > 0) {
                    let headerErrors = [];                
                    for (let i = 0; i < csvErrors.length; i++) {
                        let errorMsg = csvErrors[i].message;                    
                        if (errorMsg.startsWith("Header name ")) {                        
                            headerErrors.push(errorMsg);
                        }  
                        else {
                            errorMessage.push(errorMsg);
                        }                      
                    }                    
                    if (headerErrors.length > 0) {
                        errorMessage = headerErrors;
                    }                           
                }      
            }           
        }
        return errorMessage;
    }


}
