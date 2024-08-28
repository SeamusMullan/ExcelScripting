/*
Pricing Pack Generation Version 1.1.0 (August 2024)

Written by Seamus Mullan @ Kirby Group Engineering
Last Update: 28/08/2024


Always double check the generated tables for errors or missing details


Contact
========

Email: smulan@kirbygroup.com
Intel WWID: 12277846

(Preferably don't use this unless needed)
Personal Email: seamusmullan2023@gmail.com


To-Do
======

- 

Continuous changes
===================
Note: These may have to be done manually instead
===================
	
Implementing extra changes to names (note: usually introduced by revit / modeller)


===============================================================================================
Navisworks CSV Export Legend
===============================================================================================
E -> Element
I -> Item
C -> Custom
+ -> Not from Navisworks / Created in this script
	
Required Columns in same order as Navisworks Selection Inspector
Indented sections -> all variants / occurances of a measurement
	
(E) ID
(I) GUID
(E) Category
(C) Location (or KGE_Location)
(E) Name
(E) Size
(E) Rod Length
(E) Length
	- Any other length values too (Length, length, Length 2, etc.)
(E) Unistrut Length
(+) Units (mm/No.)
(E) Angle
	- Any other angles too (Angle, angle, Angle 2, etc.)
(E) Service type
(C) Raceway Name
(C) Raceway Ref. Number 
	
===============================================================================================
	
Ideally,the csv we are importing should have all of the data in the order shown here,
excluding any columns that get created in the script (marked with '+')
	
Repeated data should be merged into a single column and placed according to the above legend
	
===============================================================================================
*/

function log(msg: string) {
	let loggingOn = 1;
	if (loggingOn === 1) {
		console.log(msg);
	}
}

function removeControlCharacters(str: string): string {
	// This regex matches all control characters
	const controlCharsRegex = /[\u0000-\u001F\u007F-\u009F]/g;
	return str.replace(controlCharsRegex, '');
}

let nameIndex = 0;
let unistrutIndex = 0;
let angleIndex = 0;

function updateHeaders(headerValues: (string | number | boolean)[]) {
	nameIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("name")
	);
	unistrutIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("unistrut")
	);
	angleIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("angle")
	);
}

function main(workbook: ExcelScript.Workbook) {
	// Get the active worksheet
	let sheet = workbook.getActiveWorksheet();
	let usedRange = sheet.getUsedRange();

	// Optimize header processing
	let headers = usedRange.getRow(0);
	let headerValues = headers.getValues()[0];

	// Process headers once
	headerValues = headerValues.map(header => {
		if (typeof header === 'string') {
			return header
				.replace(/Element|Item|Custom|KGE_/gi, "")
				.replace("\n", "")
				.replace("\\r\\n", "");
		}
		return header.toString();
	});
	headers.setValues([headerValues]);

	// Set row height and apply autofilter
	headers.getFormat().setRowHeight(25);
	sheet.getAutoFilter().apply(headers);
	sheet.getFreezePanes().freezeRows(1);


	// Find important column indices
	let nameIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("name")
	);
	let unistrutIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("unistrut")
	);
	let angleIndex = headerValues.findIndex(header =>
		typeof header === 'string' && header.toLowerCase().includes("angle")
	);

	if (nameIndex === -1 || unistrutIndex === -1 || angleIndex === -1) {
		throw new Error("Required columns not found");
	}

	// Process data in bulk
	let values = usedRange.getValues();

	// rename angle header to just Angle
	values[0][angleIndex] = "Angle";

	log(values[0].toString());

	// Update names
	/*
	
	NOTE: These are specific for some parts (M12 -> M10) and the list below should be double checked often to keep it updated!

	*/

	const nameMap: { key: string, value: string }[] = [
		{ key: 'P1428-H-', value: 'M1116' },
		{ key: 'M12', value: 'M10' },
		{ key: 'P1062', value: 'P1020' }
	];
	for (let i = 1; i < values.length; i++) {
		let name = values[i][nameIndex].toString();

		nameMap.forEach((item) => {
			name.replace(item.key, item.value);
		});

		// if the name is in the form [words] [number], remove the number
		// a.k.a, only remove number when seperated by a space, eg:
		// Casework 2 -> Casework
		// MIDAS_Plate_2 -> MIDAS_Plate_2
		const parts = name.split(" ");
		if (parts.length > 1 && /^\d+$/.test(parts[parts.length - 1])) {
			name = parts.slice(0, -1).join(" ");
		}

		values[i][nameIndex] = name;
	}

	// Process lengths and angles
	for (let col = 0; col < headerValues.length; col++) {
		const header = headerValues[col];
		//log("On header " + header);
		if (header.toString().toLowerCase().match("length")) {
			//log("Header " + header + " is a length header");
			for (let row = 1; row < values.length; row++) {
				values[row][col] = values[row][col].toString().replace(" mm", "");
				let value = values[row][col];
				if (value > 1) {
					values[row][unistrutIndex] = value.toString().replace(" mm", "");
					//log("Value " + value + " added to unistrut column on row " + row + " from column " + header);
				}
			}
		} else if (header.toString().toLowerCase().match("angle") && col !== angleIndex) {
			for (let row = 1; row < values.length; row++) {
				let value = values[row][col];
				if ((value > 1) || (value.toString().toLowerCase().match("°"))) {
					values[row][angleIndex] = value;
				}
			}
		}


	}

	// Fill blank cells in Unistrut column with 1
	for (let i = 1; i < values.length; i++) {
		if (values[i][unistrutIndex] === "") {
			values[i][unistrutIndex] = 1;
		}
	}

	// Write processed data back to the sheet
	usedRange.setValues(values);

	// Create table
	let table = workbook.addTable(usedRange, true);

	// Add "Units (mm/No.)" column

	// remove any length column (except unistrut)
	// for (let col = 0; col < headerValues.length; col++) {
	// 	let header = headerValues[col];
	// 	header = removeControlCharacters(header.toString());
	// 	if (header.toString().toLowerCase().match("length") && !header.toString().toLowerCase().match("unistrut")) {
	// 		let column = table.getColumnByName(header);
	// 		try {
	// 			column.delete()
	// 			col--;
	// 		} catch (error) {
	// 			log(error);
	// 		}
	// 	}
	// }

	if (nameIndex === -1 || unistrutIndex === -1 || angleIndex === -1) {
		throw new Error("Required columns not found");
	}

	values[0][angleIndex] = "_Angle";

	let tableRange = workbook.getActiveWorksheet().getUsedRange();

	// delete all length columns (except unistrut)
	let colCount = tableRange.getColumnCount();
	for (let i = 0; i < colCount; i++) {
		let cell = sheet.getCell(0, i);
		let cellVal = cell.getValue().toString();
		if (cellVal != "") {
			let col: ExcelScript.TableColumn = table.getColumnByName(cellVal.toLowerCase());
			log("checking column" + i);
			log("column" + i + " named: " + cellVal);
			col.setName(cellVal.replace("\n", ""));

			// match headers with length in the name, select and delete their column (if not unistrut)
			if (cellVal.match("[Ll][Ee][Nn][Gg][Tt][Hh].*")) {
				if (!(cellVal.match("[Uu][Nn][Ii][Ss][Tt][Rr][Uu][Tt]"))) {
					log("Deleting column: " + cellVal);
					col.delete();
					updateHeaders(headerValues);
					i--;
				}
			} else if (cellVal.match("[Aa][Nn][Gg][Ll][Ee].*") && !cellVal.match("_")) {
				log("Deleting column: " + cellVal);
				log(angleIndex.toString());
				col.delete();
				updateHeaders(headerValues);
				i--;
			}
		}
	}



	// update important column indices
	updateHeaders(headerValues);


	let angleHeader = sheet.getRange("1:1").find("_Angle", {
		completeMatch: false,
		matchCase: false,
		searchDirection: ExcelScript.SearchDirection.forward
	});
	let newAngleIndex = angleHeader.getColumnIndex();

	let unitsIndex = table.addColumn(newAngleIndex, undefined, "Units (mm/No.)").getIndex();

	if (nameIndex === -1 || unistrutIndex === -1 || angleIndex === -1) {
		throw new Error("Required columns not found");
	}
	// updateHeaders(headerValues);
	// log("Angle Header: " + angleIndex.toString());
	// values[0][angleIndex] = "Angle";

	tableRange = workbook.getActiveWorksheet().getUsedRange();

	// Set formula for Units column
	let unitsFormula = `=IF(INDIRECT(ADDRESS(ROW(),COLUMN()-1))>1,"mm","No.")`;
	// log(unitsIndex.toString());
	// for (let i = 1; i < values.length; i++) {
	// 	values[i][unitsIndex] = unitsFormula;
	// }

	let unitsColumn = table.getColumnByName("Units (mm/No.)");
	unitsColumn.getRangeBetweenHeaderAndTotal().setFormula(unitsFormula);
	// Rename worksheet
	sheet.setName("Quantifications");

	updateHeaders(headerValues);
	log("Angle Header: " + angleIndex.toString());
	log(values[0][angleIndex].toString());
	values[0][angleIndex] = "Angle";


	// Remove Clearance Zones
	let rowCount = tableRange.getRowCount();
	for (let i = 1; i < rowCount; i = i + 1) {
		let cell = sheet.getCell(i, nameIndex);
		let cellVal = cell.getValue().toString();
		if (cellVal.toLowerCase().match("clearancezone.*")) {
			let row = table.getRangeBetweenHeaderAndTotal().getRow(i - 1);
			log("Deleting row " + i + " with Element ID" + row.getColumn(0).getValue())
			row.getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
			i--;
		}
	}

	// Format columns
	let format = table.getRange().getFormat();
	format.autofitColumns();
	format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	format.setIndentLevel(0);
	format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

	// Create Pivot Table
	let pivotSheet = workbook.addWorksheet("Pivot Table");
	let pivotTable = pivotSheet.addPivotTable("PivotTable1", table.getRange(), "A3");

	// Configure Pivot Table
	let unitsHierarchy = pivotTable.getHierarchy("Units (mm/No.)");
	if (unitsHierarchy) {
		pivotTable.addColumnHierarchy(unitsHierarchy);
	}
	pivotTable.getLayout().setShowRowGrandTotals(false);
}