var Excel = require('exceljs');

var workbook = new Excel.Workbook();





console.log("App is running");

var filename = "../ExcelFiles/file1.xlsx";
workbook.xlsx.readFile(filename).then(function() {

	var outPutSheet = workbook.addWorksheet('Output');
        // use workbook

        var mcc_fees = workbook.getWorksheet('MCC & FEES');

        var tid_mcc = workbook.getWorksheet('TID & MCC That Applies');

        var txn_report = workbook.getWorksheet('Sample of Txn Report');

        console.log(mcc_fees.getCell('D2').value);

        var mcc = {};
        var tmcc = {};



        mcc_fees.eachRow(function(row, rowNumber) {
            //console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

            if(rowNumber > 1) {
            	mcc[row.values[1]] = [
            		row.values[4], row.values[5], row.values[6]
            	];
            }




            
        });

        mcc_fees = null;

        delete mcc_fees;

         tid_mcc.eachRow(function(row, rowNumber) {
            //console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

            if(rowNumber > 1) {
            	tmcc[row.values[1]] = [
            		row.values[6]
            	];
            }


            
        });

          tid_mcc = null;

          delete tid_mcc;

         

         txn_report.eachRow(function(row, rowNumber) {
            //console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

            if(rowNumber > 1) {

            	thisTid = row.values[4];

            	thisAmount = row.values[6];

            	

            	if(tmcc[thisTid] != undefined) {
            		thisMcc = tmcc[thisTid][0];

            		if(mcc[thisMcc] != undefined) {
            			thisDiscount = mcc[thisMcc][0];

            			thisAmountCap = mcc[thisMcc][1];

            			thisFeeCap = mcc[thisMcc][2];

            			if(thisAmount >= thisAmountCap && thisAmountCap != 0){
            				thisMsc = thisFeeCap/4;
            			}
            			else {
            				thisMsc = (thisDiscount * thisAmount)/4;
            			}

            			outPutRow = [];

            			for(i = 1; i <= row.values.length; i++) {
            				outPutRow.push(row.values[i]);
            			}

            			outPutRow.push(thisMcc);
            			outPutRow.push(thisMsc);

            			outPutSheet.addRow(outPutRow);


            				//console.log(outPutRow);
  

            			
            		}

            	



            	}

            	

            }
            else {
            	var outPutHeaders = [];

            	for(i = 1; i <= row.values.length; i++){
            		//outPutHeader = {header: row.values[i], key: row.values[i]};
            		outPutHeaders.push(row.values[i]);
            	}
            	outPutHeaders.push("MCC");
            	outPutHeaders.push("MSC");

            	outPutSheet.addRow(outPutHeaders);


            	
            }

            
        });


         	var outPutHeaders = [];


            	outPutHeaders.push("MCC");
            	outPutHeaders.push("MSC");

            	outPutSheet.addRow(outPutHeaders);


console.log("finished bomb read through");
       // console.log(mcc[16]);

       // console.log(tmcc["2011177Y"]);

       // write to a file
       //var workbooknew = createAndFillWorkbook();
       workbook.xlsx.writeFile("../ExcelFiles/newfile1.xlsx")
           .then(function() {
               // done
               console.log("finished update");
           });
    });

