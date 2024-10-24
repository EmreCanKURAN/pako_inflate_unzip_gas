
# ZIP Extractor for Google Apps Script

## Overview
This project provides a ZIP extraction tool using Google Apps Script that extracts the contents of a ZIP file stored on Google Drive. It parses the ZIP file and extracts all files, maintaining folder structure and decompressing files if necessary, using the `pako` library for handling Deflate compression.

## Features
- **MS-DOS Timestamp Decoding**: Decodes file timestamps stored in MS-DOS format.
- **File and Folder Structure**: Preserves the folder structure of the ZIP file while extracting.
- **Custom Turkish Character Decoding**: Handles Turkish characters via a custom character mapping for older encodings.
- **Decompression with Pako**: Uses the `pako` library to decompress files stored with Deflate compression.
- **MIME Type Handling**: Automatically detects and assigns the correct MIME type for files based on their extensions.

## Requirements
- Google Apps Script environment
- ZIP file stored on Google Drive
- `pako.min.gs` library (for Deflate decompression)

## Installation
1. **Open Google Apps Script Editor**: Create a new Google Apps Script project by navigating to https://script.google.com/.
2. **Copy and Paste the Script**: Copy the provided JavaScript code for `ZipExtractor` and `extractZip` functions into the editor.
3. **Add Pako Library**:
    - Download the `pako.min.gs`.
    - Include `pako.min.gs` in your Apps Script project by creating a new file in the script editor and pasting the minified code.
4. **Save the Project**.

## Usage
1. **Modify the `main()` function**:
    - Set the `zipFileId` variable to the Google Drive ID of the ZIP file you want to extract.
    - Optionally, modify the `rootFolderName` to customize the destination folder name.

2. **Run the Script**:
    - In the script editor, click on the `Run` button to execute the script. The script will automatically extract the ZIP file into a folder in your Google Drive.

### Example

```javascript
function main() {
  const zipFileId = 'your-zip-file-id-here'; // Replace with your ZIP file ID

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
```

## Code Breakdown

### `ZipExtractor` Class
- **Constructor**: Creates a folder in Google Drive to extract the contents of the ZIP file. It checks if the folder already exists to avoid duplicates.
- **MS-DOS Time Decoding**: The `decodeDosTime` function translates timestamps from MS-DOS format to a readable date-time format.
- **Custom Turkish Mapping**: Handles file names using non-standard characters, specifically Turkish characters, and maps them to their correct Unicode equivalents.
- **Folder Creation**: The `createFolderStructure` method ensures that the extracted folder structure matches the original ZIP file.
- **File Extraction**: Extracts both compressed and non-compressed files using the `pako` library for Deflate compression.

### `extractZipFile` Method
The `extractZipFile` method starts the ZIP extraction process by reading the ZIP file from Google Drive and initiating the recursive header parsing to extract files.

### `extractZip` Function
This is a helper function that can be called from outside the class, which takes a ZIP file ID and a folder name, and extracts the contents of the ZIP file into that folder.

## Troubleshooting
- **Pako Decompression Issues**: If Deflate-compressed files fail to decompress, check if the `pako.min.gs` file is correctly added to the Project or zip file compression method is deflate.
- **Folder Duplication**: If folders are duplicated, ensure that the folder name is properly sanitized, and existing folders are checked before creating new ones.

## License
This project is licensed under the MIT License. See the LICENSE file for details.
