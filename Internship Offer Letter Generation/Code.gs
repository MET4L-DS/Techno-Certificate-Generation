function generateInternshipOfferLetters() {
	const SPREADSHEET_ID = "10tV2rnVgtDHYqmjQpqahbNENb9-JLAK-ww85Ls8yhu4";
	const TEMPLATE_FILE_ID = "1103jh9dmRFInkybV1WICRgwg0aDUYiegJ95wUp-0rCA";
	const DESTINATION_FOLDER_ID = "1JCHJ3ABAMA08kE6s0ANXrQ4KNChyt3N1";

	const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
	const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

	// Use openById if standalone, or getActiveSpreadsheet() if bound to the sheet
	const students = readSheetRequest(SPREADSHEET_ID);

	students.forEach((student) => {
		createInternshipOfferLetter(student, templateFile, destinationFolder);
	});
}

function createInternshipOfferLetter(data, templateFile, folder) {
	// Create a temp copy of the template
	const tempFile = templateFile.makeCopy(
		`${data["STUDENT_NAME"]} - Internship Offer Letter`,
		folder
	);
	const tempDoc = DocumentApp.openById(tempFile.getId());
	const body = tempDoc.getBody();

	// Replace placeholders dynamically based on keys in data
	// This expects the placeholder in Doc to match Header name exactly: {{STUDENT_NAME}}
	Object.keys(data).forEach((key) => {
		body.replaceText(`{{${key}}}`, data[key]);
	});

	// Convert to PDF
	tempDoc.saveAndClose();
	const pdfName = `${data["STUDENT_NAME"]} - Internship Offer Letter.pdf`;
	const pdfBlob = tempFile.getAs(MimeType.PDF);
	pdfBlob.setName(pdfName);

	// Check for existing files and replace them (Overwrite logic)
	const existingFiles = folder.getFilesByName(pdfName);
	while (existingFiles.hasNext()) {
		existingFiles.next().setTrashed(true);
	}

	folder.createFile(pdfBlob);
	tempFile.setTrashed(true);

	// Send Email
	if (data.EMAIL) {
		const subject = `Internship Offer Letter - ${data.STUDENT_NAME}`;
		const message = `Dear ${data.STUDENT_NAME},\n\nPlease find attached your Internship Offer Letter.\n\nRegards,\nTechnoMedia Software Solutions Pvt Ltd`;

		GmailApp.sendEmail(data.EMAIL, subject, message, {
			attachments: [pdfBlob],
			name: "TechnoMedia HR",
		});
	}
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
