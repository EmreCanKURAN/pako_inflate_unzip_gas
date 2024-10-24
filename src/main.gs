function main() {
  const zipFileId = '12jFnvBR2KisA9fzIH1kL57Jutf8uSIr3'; // Replace with your ZIP file ID

  // Get the file object using the file ID
  const file = DriveApp.getFileById(zipFileId);

  // Get the name of the file
  const rootFolderName = file.getName();

  // Optionally, remove the .zip extension from the folder name
  const folderNameWithoutExtension = rootFolderName.replace(/\.zip$/i, '');

  // Explicitly create an instance of the ZipExtractor class
  const zipExtractor = new ZipExtractor(folderNameWithoutExtension);

  // Call the extractZipFile method of ZipExtractor to start extraction
  zipExtractor.extractZipFile(zipFileId);
}