function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page');
}

// Checks if the EmployeeID exists in the sheet and returns details including BusNumber and RoomNumber
function checkEmployeeId(employeeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

  // Get header row to find column indices for BusNumber and RoomNumber
  const headers = data[0];
  const busIdx = headers.indexOf('BusNumber');
  const roomIdx = headers.indexOf('RoomNumber');
  const busSPOCIdx = headers.indexOf('BusSPOCNumber');
  const nameIdx = 1; // Name column index (assuming it's column B)

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] + '').toLowerCase() === (employeeId + '').toLowerCase()) {
      const collected = data[i][2];
      const time = data[i][3];
      const timeStr = time
        ? Utilities.formatDate(new Date(time), tz, 'dd-MMM-yyyy hh:mm a')
        : '';
      
      // Get room partners - find all employees with same room number excluding current employee
      const currentRoom = roomIdx !== -1 ? data[i][roomIdx] : '';
      const currentName = data[i][nameIdx];
      const roomPartners = [];
      
      if (currentRoom && currentRoom !== '') {
        for (let j = 1; j < data.length; j++) {
          if (j !== i && // Not the current employee
              roomIdx !== -1 && 
              data[j][roomIdx] === currentRoom && // Same room number
              data[j][nameIdx] && // Has a name
              data[j][nameIdx] !== '') { // Name is not empty
            roomPartners.push(data[j][nameIdx]);
          }
        }
      }
      
      return {
        row: i + 1,
        name: data[i][1],
        collected: collected,
        timeStr: timeStr,
        busNumber: busIdx !== -1 ? data[i][busIdx] : '',
        roomNumber: roomIdx !== -1 ? data[i][roomIdx] : '',
        employeeId: employeeId,
        busSPOCNumber: busSPOCIdx !== -1 ? data[i][busSPOCIdx] : '',
        roomPartners: roomPartners
      };
    }
  }
  return null;
}


function updateStatus(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const current = sheet.getRange(row, 3).getValue();
  if (current === 'Yes') {
    return "Already marked ✅ You cannot update again.";
  }
  sheet.getRange(row, 3).setValue('Yes');                // Shirt Collected
  sheet.getRange(row, 4).setValue(new Date());           // Timestamp
  return 'Updated Successfully ✅';
}

// Handle ID proof upload from HTML
function uploadIdProof(data) {
  // data: { employeeId, idProofType, fileName, mimeType, base64 }
  try {
    var folderName = 'ID Proofs';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var fileBytes = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(fileBytes, data.mimeType, data.fileName);
    var file = folder.createFile(blob);
    file.setName(data.fileName);

    // Optionally, you can log or update the sheet with the file URL
    // For example, find the employee row and set the file URL in a column
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange().getValues();
    var headers = dataRange[0];
    var idCol = 0; // Assuming EmployeeID is column A
    var idTypeCol = headers.indexOf('IDProof');
    if (idTypeCol === -1) {
      sheet.insertColumnAfter(headers.length);
      sheet.getRange(1, headers.length + 1).setValue('IDProof');
      idTypeCol = headers.length;
    }
    var urlCol = headers.indexOf('IDProofURL');
    if (urlCol === -1) {
      sheet.insertColumnAfter(headers.length + 1);
      sheet.getRange(1, headers.length + 2).setValue('IDProofURL');
      urlCol = headers.length + 1;
    }
    for (var i = 1; i < dataRange.length; i++) {
      if ((dataRange[i][idCol] + '').toLowerCase() === (data.employeeId + '').toLowerCase()) {
        sheet.getRange(i + 1, idTypeCol + 1).setValue(data.idProofType);
        sheet.getRange(i + 1, urlCol + 1).setValue(file.getUrl());
        break;
      }
    }
    return 'ID Proof uploaded successfully!';
  } catch (e) {
    return 'Error uploading file: ' + e.message;
  }
}
