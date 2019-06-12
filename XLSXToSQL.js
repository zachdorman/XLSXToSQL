const xlsx = require('node-xlsx');
const fs = require('fs');

let FileToArray = {
    XLSX : function(srcFile, hasHeader){

        let obj = xlsx.parse(srcFile); // parses a file
        let returnArray = new Array();
    
        //looping through all sheets    
        for(let i = 0; i < obj.length; i++)
        {            
            let rows = new Array();    
            let sheet = obj[i];

            let headerRow = new Array();
            if(hasHeader){
                //Grab the Header Row in the First Position
                headerRow = sheet['data'][0];
            }
            else{
                //Header Doesn't exist. Generate generic Column Names                
                for(let r = 0; r < sheet['data'].length; r++){
                    headerRow.push(`Column${r}`);
                }
            }
            
            //loop through all rows in the sheet
            //Ignore the Header Row
            for(let j = 1; j < sheet['data'].length; j++)
            {
                    //add the row to the rows array
                    rows.push(sheet['data'][j]);
            }

            returnArray.push({Header: headerRow, Rows: rows});
        }
        
        return returnArray;        
    }
}

//Expects
//{Header: obj, Rows: obj}    
//Functions that process the input into the desired output
let ArrayToSQL = {
    
    ToSELECT: function(inputArrays){
        
        //Use to check to see if we are at the end of rows or columns to determine when NOT to include ','s or 'UNION ALL's
        const isInLastPosition = function(currentArrayNumber, totalArrayLength){

            if(currentArrayNumber == totalArrayLength-1){
                return true;
            }
            else{
                return false;
            }
        }

        //TODO: Input is fed in with the ability to do multiple sheets in 1. But at this time I have no need for it. 
        //So just grab the first sheet
        let inputArray = inputArrays[0];


        let allRowsAsSelects = "";
        for (let i_datarows = 0; i_datarows < inputArray.Rows.length; i_datarows++) {
            const currentRow = inputArray.Rows[i_datarows];

            //Clean the inputArray.Rows of empty rows
            if(currentRow.length == 0){
                continue;
            }

            
            let currentSelect = 'SELECT ';

            for (let i_datacolumn = 0; i_datacolumn < currentRow.length; i_datacolumn++) {
                const currentColumnValue = currentRow[i_datacolumn];
                //Add the current data column and alias it as the header row in the same position
                //Scrub special characters out of the Header Title 
                currentSelect += `'${currentColumnValue}' as ${inputArray.Header[i_datacolumn].replace(/[^a-zA-Z]/g,"")}`;

                if(isInLastPosition(i_datacolumn,currentRow.length) == false){
                    currentSelect += ',';
                }
            }
            
            //Do Not Add a Union All Statement on the last SELECT
            if(isInLastPosition(i_datarows,inputArray.length)){
                currentSelect += '\n';
            }
            else{
                currentSelect += ' UNION ALL\n';
            }

            allRowsAsSelects += currentSelect;
        }                     
        
        //Check that the END of the SELECT statement doesn't contain a UNION ALL with no further rows
        if(allRowsAsSelects.substr(allRowsAsSelects.length-10,allRowsAsSelects.length-2).trimRight() == "UNION ALL"){
            allRowsAsSelects = allRowsAsSelects.substr(0, allRowsAsSelects.length-11);
        }
        return allRowsAsSelects;
    }

}

//Input Array For Processing
let FilesToETL = [
    {
        FileLocation : "48,750 FLEET190529.xlsx",        
        OutputFileName: "48,750 FLEET190529.sql",
        ProcessingFunction: ArrayToSQL.ToSELECT,
        HasHeader: true
    },
    {
        FileLocation : "771,044 FLEET190508.xlsx",        
        OutputFileName: "771,044 FLEET190508.sql",
        ProcessingFunction: ArrayToSQL.ToSELECT,
        HasHeader: true
    },
];

//Loop through input array and run the desired processing function as determined in the Array
FilesToETL.forEach((fileInformation) => {
    
    let fileAsArray = FileToArray.XLSX(fileInformation.FileLocation,true);    
    let sqlReturn = fileInformation.ProcessingFunction(fileAsArray, fileInformation.HasHeader);
    fs.writeFileSync(fileInformation.OutputFileName,sqlReturn);
});