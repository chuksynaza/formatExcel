exports.myExcelFormatter = function (fileToProcess, fileName, whereToProcess, formatResponse) {

	formatResponse("Started Parsing");

	

	inputFile = fileToProcess;

	let d = new Date();

	let theDate =  d.getFullYear() + "-" + d.getMonth() + "-" + d.getDate() + " " + d.getHours() + "-" + d.getMinutes() + "-" + d.getSeconds();

	fileName = fileName.slice(0, fileName.length - 5);
	
	outPutFile = whereToProcess + fileName + " " + theDate + " output.xlsx";

	formatResponse("The expected output file: " + outPutFile);

	let Excel = require('exceljs');

	let workbook = new Excel.Workbook();

	let workbooknew = new Excel.Workbook();

	let outPutSheet = workbooknew.addWorksheet('Output');

	let errorLogSheet = workbooknew.addWorksheet('Error Log');

	outPutSheet = workbooknew.getWorksheet('Output');

	errorLogSheet = workbooknew.getWorksheet('Error Log');

	errorLogSheet.addRow(["What Happened", "Who Caused It", "Where Did it Happen", "Amount"]);

	formatResponse("App is running");

	workbook.xlsx.readFile(inputFile).then(function() {

		formatResponse("processing");

		let mcc_fees = workbook.getWorksheet('MCC & FEES');

		let tid_mcc = workbook.getWorksheet('TID & MCC That Applies');

		let txn_report = workbook.getWorksheet('Sample of Txn Report');

		formatResponse(mcc_fees.getCell('D2').value);

		let mcc = {};
		let tmcc = {};

		mcc_fees.eachRow(function(row, rowNumber) {
	            //formatResponse('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

	            if(rowNumber > 1) {
	            	mcc[row.values[1]] = [
	            	row.values[4], row.values[5], row.values[6]
	            	];
	            }

	        });

		delete mcc_fees;

		tid_mcc.eachRow(function(row, rowNumber) {
	            //formatResponse('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

	            if(rowNumber > 1) {
	            	tmcc[row.values[1]] = [
	            	row.values[6]
	            	];
	            }

	        });

		delete tid_mcc;

		formatResponse("crunching");

		txn_report.eachRow(function(row, rowNumber) {

			
	            //formatResponse('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

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

	            			

	            			}
	            			else {
	            			//Log error that the respective MCC was not found
	            			let outPutLog = ["MCC not found", thisMcc , "The MCC and FEES sheet", thisAmount];
	            			thisMcc = "MCC not Found";
	            			//thisMsc = thisAmount;
	            			thisMsc = "";
	            			errorLogSheet.addRow(outPutLog).commit();
	            		}

	            	}
	            	else {
	            		//Log error that the TID was not found

	            		let outPutLog = ["TID not found", thisTid , "The TID and MCC reference sheet", thisAmount];
	            		thisMcc = "TID not Found";
	            		//thisMsc = thisAmount;
	            		thisMsc = "";
	            		errorLogSheet.addRow(outPutLog).commit();
	            	}


	            	outPutRow = [];

	            	for(i = 1; i < row.values.length; i++) {
	            		outPutRow.push(row.values[i]);
	            	}

	            	outPutRow.push(thisMcc);
	            	outPutRow.push(thisMsc);

	            	outPutSheet.addRow(outPutRow).commit();

	            		//formatResponse(outPutRow);

	            }
	            else {
	            	let outPutHeaders = [];

	            	for(i = 1; i < row.values.length; i++){
	            		//outPutHeader = {header: row.values[i], key: row.values[i]};
	            		outPutHeaders.push(row.values[i]);
	            	}
	            	outPutHeaders.push("MCC");
	            	outPutHeaders.push("MSC");

	            	outPutSheet.addRow(outPutHeaders).commit();

	            }

	        });

		delete txn_report;

		formatResponse("finished bomb read through");

		formatResponse("committed the workbook");

	     	           formatResponse("writing");
	     	           workbooknew.xlsx.writeFile(outPutFile)
	     	           .then(function() {

	     	           	formatResponse("complete");

	     	                   formatResponse("finished update");
	     	                   formatResponse("Format Success");

	     	               });
	     	       });
};


