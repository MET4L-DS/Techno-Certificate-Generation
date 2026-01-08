// generates PDF certificates based on JSON data using a Google Doc template.
function generateCertificates() {
	const TEMPLATE_FILE_ID = "1sykS23r3RP3JUCO49eEkBG9yPHEufcvk15S7VzngvuE";
	const DESTINATION_FOLDER_ID = "1XpG0PD-pJmeLcL6Chy2EG6qHlK_iedhv";
	const INTERN_FILE_ID = "1MBd59qfjbEpHTfSXkPvBoktWdcmrEKyz";

	const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
	const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
	const internData = readJSON(INTERN_FILE_ID);

	internData.forEach((intern) => {
		createCertificate(intern, templateFile, destinationFolder);
	});
}

function createCertificate(data, templateFile, folder) {
	const tempFile = templateFile.makeCopy(data.name, folder);
	const tempDoc = DocumentApp.openById(tempFile.getId());
	const body = tempDoc.getBody();

	// replace placeholders with actual data
	body.replaceText("{{Name}}", data.name);
	body.replaceText("{{Role}}", data.role);
	body.replaceText("{{StartDate}}", data.startDate);
	body.replaceText("{{EndDate}}", data.endDate);
	body.replaceText("{{Manager}}", data.manager);
	body.replaceText("{{IssueDate}}", new Date().toLocaleDateString());

	// convert to pdf
	tempDoc.saveAndClose();
	const pdfName = `${data.name}.pdf`;
	const pdfBlob = tempFile.getAs(MimeType.PDF);
	pdfBlob.setName(pdfName);

	// delete existing file with the same name if it exists
	const existingFiles = folder.getFilesByName(pdfName);
	while (existingFiles.hasNext()) {
		existingFiles.next().setTrashed(true);
	}

	folder.createFile(pdfBlob);
	tempFile.setTrashed(true);
}

function readJSON(fileId) {
	const file = DriveApp.getFileById(fileId);
	const content = file.getBlob().getDataAsString();
	return JSON.parse(content);
}
