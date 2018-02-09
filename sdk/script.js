formatExcel = require("./excelFormatter");
let wait = false;
function formatFile(fileToProcess, fileName, whereToProcess, formatExcel){

	function formatResponse(message){
		console.log(message);
		if(message == "reading"){
			document.getElementById("statusupdate").innerHTML = "Status: Reading workbook...";
		}
		else if(message == "processing"){
			document.getElementById("statusupdate").innerHTML = "Status: Processing workbook, please wait...";
		}
		else if(message == "computing") {
			document.getElementById("statusupdate").innerHTML = "Status: Crunching numbers...";
		}
		else if(message == "writing"){
			document.getElementById("statusupdate").innerHTML = "Status: Writing new workbook...";
		}
		else if(message == "complete") {
			wait = false;
			document.getElementById("processData").style.backgroundColor = "#3a4554";
			document.getElementById("processData").style.cursor = "pointer";
			document.getElementById("statusupdate").innerHTML = "Status: Completed, please check for the new workbook";
		}

	}

	formatResponse("reading");

	console.log("Started Processing");

	formatExcel.myExcelFormatter(fileToProcess, fileName, whereToProcess, formatResponse);
}

document.getElementById("processData").onclick = function() {

	if(wait) {
		return false;
	}

	wait = true;
	document.getElementById("processData").style.backgroundColor = "#696b6e";
	document.getElementById("processData").style.cursor = "default";

	let inputFile = document.getElementById("inputfile");
	let inputName = inputFile.files[0].name;
	let fileP = inputFile.value.slice(0, inputFile.value.length - inputName.length);
	//let outputFolder = document.getElementById("outputfolder");
	//alert(inputFile.value);
	//alert(fileP);
	//alert(outputFolder.value);

	formatFile(inputFile.value, inputName, fileP, formatExcel);

}