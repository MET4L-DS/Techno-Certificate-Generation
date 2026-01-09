function generateAppointmentLetters() {
	const SPREADSHEET_ID = "10ynPS_Yzc_i2igHbHCQa18iZrbaz7MIC_M7yibF03X8"; // Replace with your actual Spreadsheet ID
	const TEMPLATE_FILE_ID = "13rACPYjJQqyVGw6qdRGw2YDe9cprJ97o-LdfpNxDzY0";
	const DESTINATION_FOLDER_ID = "1LX-hbAhU9GkN9B9VMBO4U3Uf1Wjwbp31";

	const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
	const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

	// Use openById if standalone, or getActiveSpreadsheet() if bound to the sheet
	const employees = readSheetRequest(SPREADSHEET_ID);

	employees.forEach((employee) => {
		createAppointmentLetter(employee, templateFile, destinationFolder);
	});
}

function createAppointmentLetter(data, templateFile, folder) {
	// Create a temp copy of the template
	const tempFile = templateFile.makeCopy(
		`${data["EMP_NAME"]} - Appointment Letter`,
		folder
	);
	const tempDoc = DocumentApp.openById(tempFile.getId());
	const body = tempDoc.getBody();

	// Replace placeholders dynamically based on keys in data
	// This expects the placeholder in Doc to match Header name exactly: {{EMP_NAME}}
	Object.keys(data).forEach((key) => {
		body.replaceText(`{{${key}}}`, data[key]);
	});

	// Convert to PDF
	tempDoc.saveAndClose();
	const pdfName = `${data["EMP_NAME"]} - Appointment Letter.pdf`;
	const pdfBlob = tempFile.getAs(MimeType.PDF);
	pdfBlob.setName(pdfName);

	// Check for existing files and replace them (Overwrite logic)
	const existingFiles = folder.getFilesByName(pdfName);
	while (existingFiles.hasNext()) {
		existingFiles.next().setTrashed(true);
	}

	folder.createFile(pdfBlob);
	tempFile.setTrashed(true);
}

function readSheetRequest(spreadsheetId) {
	// If the script is bound to the sheet, use: SpreadsheetApp.getActiveSpreadsheet();
	// Otherwise use openById:
	let sheet;
	try {
		sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
	} catch (e) {
		// Fallback if ID is invalid or script is bound
		sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
	}

	const data = sheet.getDataRange().getDisplayValues();
	const headers = data[0];
	const rows = data.slice(1);

	return rows.map((row) => {
		const obj = {};
		headers.forEach((header, index) => {
			obj[header] = row[index];
		});
		return obj;
	});
}
