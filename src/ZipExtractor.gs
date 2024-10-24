// Define the ZipExtractor class
class ZipExtractor {
  constructor(rootFolderName) {
    this.rootFolder = DriveApp.createFolder(rootFolderName);
    this.existingFolders = {};
  }

  // Function to decode MS-DOS time and date into a readable format
  decodeDosTime(zipBytes, offset) {
    const dataView = new DataView(zipBytes.buffer, offset, 4);
    const time = dataView.getUint16(0, true);
    const date = dataView.getUint16(2, true);

    const seconds = (time & 0x1F) * 2; // Seconds divided by 2
    const minutes = (time >> 5) & 0x3F; // Bits 5-10: minute
    const hours = (time >> 11) & 0x1F; // Bits 11-15: hour

    const day = date & 0x1F; // Bits 0-4: day
    const month = (date >> 5) & 0x0F; // Bits 5-8: month
    const year = ((date >> 9) & 0x7F) + 1980; // Bits 9-15: years from 1980

    return { seconds, minutes, hours, day, month, year };
  }

  // Function to get the MIME type based on the file extension
  getMimeType(fileName) {
    const extension = fileName.split('.').pop().toLowerCase();
    const mimeTypes = {
      pdf: 'application/pdf',
      docx:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      xlsx:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      zip: 'application/zip',
      txt: 'text/plain',
      // Add more extensions and MIME types as needed
    };
    return mimeTypes[extension] || 'application/octet-stream'; // Default to binary stream
  }

  // Function to decode the file name based on the bit flag
  decodeFileName(fileNameBytes, bitFlag) {
    const isUtf8 = (bitFlag & (1 << 11)) !== 0; // UTF-8 bit in the General Purpose Bit Flag
    let decodedName = '';

    try {
      if (isUtf8) {
        decodedName = new TextDecoder('utf-8').decode(fileNameBytes);
      } else {
        // Use ISO-8859-9 (Turkish locale)
        decodedName = new TextDecoder('iso-8859-9').decode(fileNameBytes);
      }
    } catch (e) {
      // If decoding fails, apply custom Turkish mapping
      decodedName = this.applyCustomTurkishMapping(fileNameBytes);
    }

    return decodedName;
  }

  // Custom mapping function for Turkish characters
  applyCustomTurkishMapping(fileNameBytes) {
    const customMap = {
      0x8d: 'ı', // Turkish: ı (dotless i)
      0x9f: 'ş', // Turkish: ş
      0xa6: 'Ğ', // Turkish: Ğ
      0xa7: 'ğ', // Turkish: ğ
      0x81: 'ü', // Turkish: İ (uppercase dotted I)
      0x87: 'ç', // Turkish: ç
      0x94: 'ö', // Turkish: ö
      0x98: 'İ', // Turkish: Ş (uppercase S with cedilla)
      0x99: 'Ö', // Turkish: Ö
      0x9a: 'Ü', // Turkish: Ü (uppercase U with umlaut)
      0x9e: 'Ş', // Turkish: Ş (uppercase S with cedilla)
      0x80: 'Ç', // Turkish: Ç
    };

    // Build the new string by mapping each byte to its corresponding Turkish character
    let mappedString = '';
    for (let i = 0; i < fileNameBytes.length; i++) {
      const byteValue = fileNameBytes[i];
      if (customMap[byteValue]) {
        mappedString += customMap[byteValue];
      } else {
        mappedString += String.fromCharCode(byteValue); // Use original character if no mapping is found
      }
    }

    return mappedString;
  }

  // Function to extract the file
  extractFile(
    zipBytes,
    localHeaderOffset,
    fileName,
    compressedSize,
    uncompressedSize,
    compressionMethod,
    generalPurposeFlag
  ) {
    const localHeaderSignature = [0x50, 0x4b, 0x03, 0x04];

    // Check if local file header signature is valid
    if (
      zipBytes[localHeaderOffset] !== localHeaderSignature[0] ||
      zipBytes[localHeaderOffset + 1] !== localHeaderSignature[1] ||
      zipBytes[localHeaderOffset + 2] !== localHeaderSignature[2] ||
      zipBytes[localHeaderOffset + 3] !== localHeaderSignature[3]
    ) {
      // Invalid signature, skip
      return localHeaderOffset + 4;
    }

    // Read local file header to get additional file details
    const dataView = new DataView(zipBytes.buffer, localHeaderOffset);
    const fileNameLength = dataView.getUint16(26, true);
    const extraFieldLength = dataView.getUint16(28, true);

    // Calculate the starting offset of the file data (after the local header)
    const fileDataOffset =
      localHeaderOffset + 30 + fileNameLength + extraFieldLength;

    // Create folder structure if the file is a directory
    if (fileName.endsWith('/')) {
      this.createFolderStructure(fileName);
      return fileDataOffset;
    }

    // Extract the compressed data
    const compressedData = zipBytes.subarray(
      fileDataOffset,
      fileDataOffset + compressedSize
    );

    // Check if compressedData is empty
    if (compressedData.length === 0) {
      return fileDataOffset + compressedSize;
    }

    // Decompress the data using Pako
    let fileData;
    if (compressionMethod === 8) {
      // Deflate compression
      try {
        fileData = pako.inflateRaw(compressedData);
      } catch (e) {
        // Decompression failed
        return fileDataOffset + compressedSize;
      }
    } else if (compressionMethod === 0) {
      // No compression
      fileData = compressedData;
    } else {
      // Unsupported compression method
      return fileDataOffset + compressedSize;
    }

    // Check if fileData is empty or undefined
    if (!fileData || fileData.length === 0) {
      return fileDataOffset + compressedSize;
    }

    // Save the extracted file to Google Drive
    const mimeType = this.getMimeType(fileName);
    try {
      const blob = Utilities.newBlob(
        fileData,
        mimeType,
        fileName.split('/').pop()
      );

      const parentFolder = this.createFolderStructure(
        this.getParentFolderPath(fileName)
      );
      if (parentFolder) {
        parentFolder.createFile(blob);
      }
    } catch (error) {
      // Handle error
    }

    // Return the next file header offset
    return fileDataOffset + compressedSize;
  }

  // Function to create folder structure
  createFolderStructure(folderPath) {
    if (!folderPath) {
      return this.rootFolder;
    }

    const folders = folderPath.split('/').filter((folder) => folder.length > 0);
    let currentFolder = this.rootFolder;

    for (let i = 0; i < folders.length; i++) {
      const folderName = folders[i];
      const folderKey = folders.slice(0, i + 1).join('/');

      if (this.existingFolders[folderKey]) {
        currentFolder = this.existingFolders[folderKey];
      } else {
        let existingFolder = currentFolder.getFoldersByName(folderName);
        if (existingFolder.hasNext()) {
          currentFolder = existingFolder.next();
        } else {
          currentFolder = currentFolder.createFolder(folderName);
        }
        this.existingFolders[folderKey] = currentFolder;
      }
    }
    return currentFolder;
  }

  // Function to get parent folder path
  getParentFolderPath(filePath) {
    if (!filePath.includes('/')) {
      return '';
    }
    const parts = filePath.split('/');
    parts.pop(); // Remove the file name
    return parts.join('/') + '/';
  }

  // Function to read little-endian integers from a byte array using DataView
  readLittleEndian(zipBytes, offset, length) {
    const dataView = new DataView(zipBytes.buffer, offset, length);
    if (length === 2) {
      return dataView.getUint16(0, true);
    } else if (length === 4) {
      return dataView.getUint32(0, true);
    }
    return 0;
  }

  // Recursive function to process all local file headers in the ZIP
  parseAllLocalFileHeaders(zipBytes, offset) {
    while (offset < zipBytes.length) {
      const nextOffset = this.parseLocalFileHeader(zipBytes, offset);
      if (nextOffset === -1) break; // Stop if no more headers are found or error occurs
      offset = nextOffset; // Move to the next header
    }
  }

  // Function to parse the Local File Header in the ZIP file
  parseLocalFileHeader(zipBytes, offset) {
    const signature = zipBytes.subarray(offset, offset + 4);
    if (
      signature[0] !== 0x50 ||
      signature[1] !== 0x4b ||
      signature[2] !== 0x03 ||
      signature[3] !== 0x04
    ) {
      // Invalid signature
      return -1;
    }

    const dataView = new DataView(zipBytes.buffer, offset);
    const version = dataView.getUint16(4, true);
    const flags = dataView.getUint16(6, true);
    const compressionMethod = dataView.getUint16(8, true);
    const modTime = dataView.getUint16(10, true);
    const modDate = dataView.getUint16(12, true);
    let crc32 = dataView.getUint32(14, true);
    let compressedSize = dataView.getUint32(18, true);
    let uncompressedSize = dataView.getUint32(22, true);
    const fileNameLength = dataView.getUint16(26, true);
    const extraFieldLength = dataView.getUint16(28, true);

    const modDateTime = this.decodeDosTime(zipBytes, offset + 10);

    const fileNameBytes = zipBytes.subarray(
      offset + 30,
      offset + 30 + fileNameLength
    );
    const fileName = this.decodeFileName(fileNameBytes, flags);

    // Handle data descriptor if necessary
    const dataDescriptorFlagSet = (flags & (1 << 3)) !== 0;
    if (dataDescriptorFlagSet) {
      const fileDataOffset = offset + 30 + fileNameLength + extraFieldLength;
      const endOfData = this.extractFile(
        zipBytes,
        offset,
        fileName,
        compressedSize,
        uncompressedSize,
        compressionMethod,
        flags
      );
      const dataDescriptorOffset = endOfData;
      const dataDescriptorView = new DataView(
        zipBytes.buffer,
        dataDescriptorOffset
      );
      crc32 = dataDescriptorView.getUint32(0, true);
      compressedSize = dataDescriptorView.getUint32(4, true);
      uncompressedSize = dataDescriptorView.getUint32(8, true);
      return dataDescriptorOffset + 12; // Move past data descriptor
    } else {
      // Extract and process the file
      const nextOffset = this.extractFile(
        zipBytes,
        offset,
        fileName,
        compressedSize,
        uncompressedSize,
        compressionMethod,
        flags
      );
      return nextOffset !== -1
        ? nextOffset
        : offset + 30 + fileNameLength + extraFieldLength;
    }
  }

  // Main function to initiate ZIP extraction
  extractZipFile(zipFileId) {
    const file = DriveApp.getFileById(zipFileId);
    const zipBlob = file.getBlob();
    const zipBytes = new Uint8Array(zipBlob.getBytes());

    // Start recursive extraction from the beginning (offset 0)
    this.parseAllLocalFileHeaders(zipBytes, 0);
  }
}

// Function accessible from main code to extract ZIP file
function extractZip(zipFileId, rootFolderName) {
  const zipExtractor = new ZipExtractor(rootFolderName);
  zipExtractor.extractZipFile(zipFileId);
}
