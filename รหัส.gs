// === GLOBAL CONSTANTS ===
const SLIDE_TEMPLATE_ID = '1AAO4FDlxOykqa3E1E621U7zjE-7Zvl6JIUYIMDx32iM';
const OUTPUT_FOLDER_ID = '1WBD--l_fQtRidTHTHRl_n336FVnl4YzT';
const REPORTS_SHEET_NAME = 'รายงานผลการพัฒนาตนเอง';
const PERSONNEL_SHEET_NAME = 'บุคลากร';
const SETTINGS_SHEET_NAME = 'ตั้งค่าระบบ'; // หรือใช้ PropertiesService

const DEFAULT_WEBSITE_TITLE = "ระบบรายงานผลการพัฒนาตนเอง โรงเรียนบ้านนานวล";
const DEFAULT_FAVICON_URL = 'https://img5.pic.in.th/file/secure-sv1/273218374_306049724897300_8948544915894559738_n.png';
const DEFAULT_LOGO_URL = 'https://img5.pic.in.th/file/secure-sv1/273218374_306049724897300_8948544915894559738_n.png';
const DEFAULT_FOOTER_TEXT = 'ระบบรายงานผลการพัฒนาตนเองและวิชาชีพ โรงเรียนบ้านนานวล';

// === INITIAL SETUP ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Setup')
    .addItem('Initialize Sheets & Folders', 'initializeApp')
    .addToUi();
}

function initializeApp() {
  ensureSheetExists(REPORTS_SHEET_NAME, ['Timestamp', 'ชื่อผู้เข้าร่วม', 'ตำแหน่ง', 'ชื่อกิจกรรม', 'วันที่เข้าร่วม', 'หน่วยงานที่จัด', 'ความรู้ที่ได้รับ', 'ผลการเข้าร่วมและการนำไปใช้', 'ลิงก์ภาพ1', 'ลิงก์ภาพ2', 'ลิงก์ลายเซ็น', 'ลิงก์PDF', 'ID']);
  ensureSheetExists(PERSONNEL_SHEET_NAME, ['ชื่อ-สกุล', 'ตำแหน่ง']);
  // ensureSheetExists(SETTINGS_SHEET_NAME, ['Key', 'Value']); // ถ้าใช้ Sheet เก็บ Config

  // Initialize PropertiesService for settings
  const properties = PropertiesService.getUserProperties();
  if (!properties.getProperty('websiteTitle')) {
    properties.setProperty('websiteTitle', DEFAULT_WEBSITE_TITLE);
  }
  if (!properties.getProperty('faviconUrl')) {
    properties.setProperty('faviconUrl', DEFAULT_FAVICON_URL);
  }
  if (!properties.getProperty('logoUrl')) {
    properties.setProperty('logoUrl', DEFAULT_LOGO_URL);
  }
  if (!properties.getProperty('footerText')) {
    properties.setProperty('footerText', DEFAULT_FOOTER_TEXT);
  }

  // Check output folder
  try {
    DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    SpreadsheetApp.getUi().alert('Initialization', 'Folders and Sheets are ready.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Output Folder ID (' + OUTPUT_FOLDER_ID + ') not found or access denied. Please check and try again.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ensureSheetExists(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1); // Freeze header row
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }
  } else {
    // Check headers if sheet exists
    if (headers && headers.length > 0) {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      let headersMatch = headers.length === currentHeaders.length;
      if (headersMatch) {
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] !== currentHeaders[i]) {
            headersMatch = false;
            break;
          }
        }
      }
      if (!headersMatch || sheet.getLastColumn() < 1) { // if no headers or mismatch
        sheet.clearContents(); // Or more careful update
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      }
    }
  }
  return sheet;
}


// === WEB APP ENTRY POINT ===
function doGet(e) {
  const properties = PropertiesService.getUserProperties();
  const title = properties.getProperty('websiteTitle') || DEFAULT_WEBSITE_TITLE;
  const faviconUrl = properties.getProperty('faviconUrl') || DEFAULT_FAVICON_URL;

  let htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
  htmlOutput.setTitle(title)
            .setFaviconUrl(faviconUrl)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return htmlOutput;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === SERVER-SIDE FUNCTIONS ACCESSIBLE FROM CLIENT ===
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getInitialData() {
  const personnel = getPersonnelList();
  const settings = getCurrentSettings();
  return {
    personnel: personnel,
    settings: settings,
    webAppUrl: getWebAppUrl()
  };
}

function getCurrentSettings() {
    const properties = PropertiesService.getUserProperties();
    return {
        websiteTitle: properties.getProperty('websiteTitle') || DEFAULT_WEBSITE_TITLE,
        logoUrl: properties.getProperty('logoUrl') || DEFAULT_LOGO_URL,
        faviconUrl: properties.getProperty('faviconUrl') || DEFAULT_FAVICON_URL,
        footerText: properties.getProperty('footerText') || DEFAULT_FOOTER_TEXT
    };
}

function getPersonnelList() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PERSONNEL_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // No data beyond header
    return data.slice(1).map(row => ({ name: row[0], position: row[1] }));
  } catch (e) {
    console.error("Error in getPersonnelList: " + e.toString());
    return [];
  }
}

function uploadFileToDrive(base64Data, fileName, mimeType, folderId) {
  try {
    const data = Utilities.base64Decode(base64Data, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(data, mimeType, fileName);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    return { id: file.getId(), url: file.getUrl(), name: file.getName() };
  } catch (e) {
    console.error("Error uploading file: " + fileName + " - " + e.toString());
    throw new Error("Failed to upload file: " + fileName + ". " + e.message);
  }
}

function processForm(formData) {
  try {
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const reportSheet = ensureSheetExists(REPORTS_SHEET_NAME, ['Timestamp', 'ชื่อผู้เข้าร่วม', 'ตำแหน่ง', 'ชื่อกิจกรรม', 'วันที่เข้าร่วม', 'หน่วยงานที่จัด', 'ความรู้ที่ได้รับ', 'ผลการเข้าร่วมและการนำไปใช้', 'ลิงก์ภาพ1', 'ลิงก์ภาพ2', 'ลิงก์ลายเซ็น', 'ลิงก์PDF', 'ID']);

    let imageUrl1 = '', imageFileId1 = '';
    let imageUrl2 = '', imageFileId2 = '';
    let signatureUrl = '', signatureFileId = '';

    // Upload image 1
    if (formData.image1 && formData.image1.base64) {
      const fileInfo1 = uploadFileToDrive(formData.image1.base64, formData.image1.name, formData.image1.type, OUTPUT_FOLDER_ID);
      imageUrl1 = fileInfo1.url;
      imageFileId1 = fileInfo1.id;
    }

    // Upload image 2
    if (formData.image2 && formData.image2.base64) {
      const fileInfo2 = uploadFileToDrive(formData.image2.base64, formData.image2.name, formData.image2.type, OUTPUT_FOLDER_ID);
      imageUrl2 = fileInfo2.url;
      imageFileId2 = fileInfo2.id;
    }

    // Upload signature
    if (formData.signature && formData.signature.base64) {
      const signatureFileInfo = uploadFileToDrive(formData.signature.base64, formData.signature.name, formData.signature.type, OUTPUT_FOLDER_ID);
      signatureUrl = signatureFileInfo.url;
      signatureFileId = signatureFileInfo.id;
    } else if (formData.signatureFile && formData.signatureFile.base64) { // Signature uploaded as file
        const signatureFileInfo = uploadFileToDrive(formData.signatureFile.base64, formData.signatureFile.name, formData.signatureFile.type, OUTPUT_FOLDER_ID);
        signatureUrl = signatureFileInfo.url;
        signatureFileId = signatureFileInfo.id;
    }


    // Create PDF
    const templateFile = DriveApp.getFileById(SLIDE_TEMPLATE_ID);
    const newFileName = 'รายงานพัฒนาตนเองของ-' + formData.name;
    const newPresentationFile = templateFile.makeCopy(newFileName, outputFolder);
    const presentation = SlidesApp.openById(newPresentationFile.getId());

    // Replace text placeholders
    presentation.replaceAllText('{{ชื่อ}}', formData.name || '');
    presentation.replaceAllText('{{ตำแหน่ง}}', formData.position || '');
    presentation.replaceAllText('{{ชื่อกิจกรรม}}', formData.activityName || '');
    presentation.replaceAllText('{{วันที่เข้าร่วมกิจกรรม}}', formData.activityDate ? new Date(formData.activityDate).toLocaleDateString('th-TH') : '');
    presentation.replaceAllText('{{หน่วยงานที่จัดกิจกรรม}}', formData.organizer || '');
    presentation.replaceAllText('{{ความรู้}}', formData.knowledge || '');
    presentation.replaceAllText('{{ผลการร่วมกิจกรรม}}', formData.results || '');

    // Replace image placeholders
    if (imageFileId1) {
      const imageBlob1 = DriveApp.getFileById(imageFileId1).getBlob();
      replaceShapeWithImage(presentation, '{{ภาพ1}}', imageBlob1);
    }
    if (imageFileId2) {
      const imageBlob2 = DriveApp.getFileById(imageFileId2).getBlob();
      replaceShapeWithImage(presentation, '{{ภาพ2}}', imageBlob2);
    }
    if (signatureFileId) {
      const signatureBlob = DriveApp.getFileById(signatureFileId).getBlob();
      replaceShapeWithImage(presentation, '{{ลายเซ็น}}', signatureBlob);
    }
    
    presentation.saveAndClose();

    // Convert to PDF
    const pdfBlob = newPresentationFile.getAs(MimeType.PDF);
    const pdfFileName = newFileName + '.pdf';
    const pdfFile = outputFolder.createFile(pdfBlob).setName(pdfFileName);
    const pdfFileLink = pdfFile.getUrl(); // Link to Drive page for the PDF
    const pdfFileViewLink = `https://drive.google.com/file/d/${pdfFile.getId()}/preview`; // For embedding/direct view


    // Save data to Google Sheet
    const newRowId = Utilities.getUuid(); // Generate a unique ID for the row
    reportSheet.appendRow([
      new Date(),
      formData.name,
      formData.position,
      formData.activityName,
      formData.activityDate ? new Date(formData.activityDate) : '',
      formData.organizer,
      formData.knowledge,
      formData.results,
      imageUrl1,
      imageUrl2,
      signatureUrl,
      pdfFileLink, // Store the Drive page URL
      newRowId
    ]);

    return { success: true, message: 'รายงานถูกสร้างและบันทึกเรียบร้อยแล้ว!', pdfLink: pdfFileLink, pdfViewLink: pdfFileViewLink, fileName: pdfFileName };

  } catch (e) {
    console.error("Error in processForm: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'เกิดข้อผิดพลาดในการประมวลผล: ' + e.message };
  }
}

function replaceShapeWithImage(presentation, shapeName, imageBlob) {
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; i++) {
    const shapes = slides[i].getShapes();
    for (let j = 0; j < shapes.length; j++) {
      const shape = shapes[j];
      if (shape.getText && shape.getText().asString().trim() === shapeName) { // Check if shape contains the placeholder text
        // The actual replacement logic can be complex to achieve "cover, crop to fill".
        // A simple replaceImage will insert the image into the shape.
        // For "cover, crop to fill", the shape itself should be configured in Google Slides
        // (Format options -> Size & Position -> Fit object in shape: Crop)
        // Then replaceImage should honor this.
        const newImage = slides[i].insertImage(imageBlob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
        shape.remove(); // Remove the old shape placeholder
        // To truly mimic "cover, crop to fill" programmatically without relying on shape settings:
        // This would require calculating aspect ratios and using image.setTransform() which is more advanced.
        // For now, we rely on the simpler insertImage and the template shape's properties.
        return; // Assume placeholder name is unique
      } else if (shape.getTitle && shape.getTitle() === shapeName) { // Alternative: check shape's title (set via Alt Text in Slides)
         const newImage = slides[i].insertImage(imageBlob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
         shape.remove();
         return;
      }
    }
  }
  console.warn("Shape with name '" + shapeName + "' not found in presentation.");
}


function getReportData(filters = {}) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORTS_SHEET_NAME);
    if (!sheet) return { data: [], personnelCount: 0, activityCount: 0 };

    let data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row

    // Apply filters
    if (filters.name) {
      data = data.filter(row => row[headers.indexOf('ชื่อผู้เข้าร่วม')] === filters.name);
    }
    if (filters.position) {
      data = data.filter(row => row[headers.indexOf('ตำแหน่ง')] === filters.position);
    }
    if (filters.year) {
      const year = parseInt(filters.year);
      data = data.filter(row => {
        const dateVal = row[headers.indexOf('วันที่เข้าร่วม')];
        return dateVal && (new Date(dateVal)).getFullYear() === year;
      });
    }
    if (filters.month) {
      const month = parseInt(filters.month) - 1; // JS months are 0-indexed
       data = data.filter(row => {
        const dateVal = row[headers.indexOf('วันที่เข้าร่วม')];
        return dateVal && (new Date(dateVal)).getMonth() === month;
      });
    }
    
    const formattedData = data.map((row, index) => {
      const dateValue = row[headers.indexOf('วันที่เข้าร่วม')];
      let formattedDate = '';
      if (dateValue) {
        try {
          formattedDate = new Date(dateValue).toLocaleDateString('th-TH', {
            year: 'numeric', month: 'long', day: 'numeric'
          });
        } catch(e){
          formattedDate = dateValue.toString(); // fallback
        }
      }
      const pdfLink = row[headers.indexOf('ลิงก์PDF')];
      const fileIdMatch = pdfLink ? pdfLink.match(/[-\w]{25,}/) : null; // Basic regex to find Drive File ID
      const pdfViewLink = fileIdMatch ? `https://drive.google.com/file/d/${fileIdMatch[0]}/preview` : '#';
      
      return {
        id: row[headers.indexOf('ID')] || (index + 1), // Use stored ID or fallback to index
        no: index + 1,
        name: row[headers.indexOf('ชื่อผู้เข้าร่วม')],
        position: row[headers.indexOf('ตำแหน่ง')],
        activityName: row[headers.indexOf('ชื่อกิจกรรม')],
        activityDate: formattedDate,
        pdfLink: pdfLink,
        pdfViewLink: pdfViewLink
      };
    });

    // Statistics
    const personnelList = getPersonnelList(); // Assuming this returns an array of objects {name: ..., position: ...}
    const personnelCount = personnelList.length;
    const activityCount = formattedData.length; // Or count unique activities if needed

    return { data: formattedData, personnelCount: personnelCount, activityCount: activityCount, headers: headers };
  } catch (e) {
    console.error("Error in getReportData: " + e.toString());
    return { data: [], personnelCount: 0, activityCount: 0, error: e.message };
  }
}


// === ADMIN FUNCTIONS ===
function verifyAdminPassword(password) {
  return password === "a123456"; // Simple password check
}

function saveAdminSettings(settings) {
  try {
    const properties = PropertiesService.getUserProperties();
    if (settings.websiteTitle) properties.setProperty('websiteTitle', settings.websiteTitle);
    if (settings.logoUrl) properties.setProperty('logoUrl', settings.logoUrl);
    if (settings.faviconUrl) properties.setProperty('faviconUrl', settings.faviconUrl);
    if (settings.footerText) properties.setProperty('footerText', settings.footerText);
    return { success: true, message: "บันทึกการตั้งค่าเรียบร้อยแล้ว" };
  } catch (e) {
    console.error("Error saving admin settings: " + e.toString());
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}

function addPersonnel(name, position) {
  try {
    if (!name || !position) {
      return { success: false, message: "กรุณากรอกชื่อและตำแหน่ง" };
    }
    const sheet = ensureSheetExists(PERSONNEL_SHEET_NAME, ['ชื่อ-สกุล', 'ตำแหน่ง']);
    // Check for duplicates
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) { // Start from 1 to skip header
        if(data[i][0] === name) {
            return { success: false, message: `ชื่อ '${name}' มีอยู่ในระบบแล้ว` };
        }
    }
    sheet.appendRow([name, position]);
    return { success: true, message: "เพิ่มข้อมูลบุคลากรเรียบร้อยแล้ว", personnel: getPersonnelList() };
  } catch (e) {
    console.error("Error adding personnel: " + e.toString());
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}

function importPersonnelFromExcel(fileData) {
  // This function expects fileData to be an array of arrays (rows of cells)
  // The client-side will need to parse the Excel file into this format (e.g., using SheetJS)
  // For a pure Apps Script solution, you'd upload the .xlsx, convert it to Google Sheets, read, then delete.
  // That's more complex. This simpler version assumes client-side parsing.
  // fileData = { name: "filename.xlsx", type: "mimetype", base64: "base64string" }

  if (!fileData || !fileData.base64) {
    return { success: false, message: "ไม่พบข้อมูลไฟล์ Excel" };
  }

  try {
    const tempFolderId = OUTPUT_FOLDER_ID; // Use the same output folder or a dedicated temp one
    const decodedData = Utilities.base64Decode(fileData.base64);
    const blob = Utilities.newBlob(decodedData, fileData.type, fileData.name);
    
    const tempFile = DriveApp.getFolderById(tempFolderId).createFile(blob);
    const spreadsheetId = tempFile.getId(); // The ID of the uploaded Excel file

    // Convert to Google Sheet format using Drive API (Advanced Service)
    // This part requires Drive API to be enabled in Services
    const convertedFile = Drive.Files.copy({ mimeType: MimeType.GOOGLE_SHEETS, title: `[Temp Import] ${fileData.name}` }, spreadsheetId);
    const tempSheetId = convertedFile.id;
    
    const ss = SpreadsheetApp.openById(tempSheetId);
    const importSheet = ss.getSheets()[0]; // Assuming data is in the first sheet
    const data = importSheet.getDataRange().getValues();
    
    DriveApp.getFileById(spreadsheetId).setTrashed(true); // Delete original Excel upload
    DriveApp.getFileById(tempSheetId).setTrashed(true); // Delete temporary Google Sheet

    if (!data || data.length <= 1) { // No data or only header
      return { success: false, message: "ไฟล์ Excel ไม่มีข้อมูล หรือมีเพียงหัวตาราง" };
    }

    const personnelSheet = ensureSheetExists(PERSONNEL_SHEET_NAME, ['ชื่อ-สกุล', 'ตำแหน่ง']);
    const existingPersonnelData = personnelSheet.getDataRange().getValues();
    const existingNames = existingPersonnelData.slice(1).map(r => r[0]); // Get existing names to avoid duplicates

    let addedCount = 0;
    let skippedCount = 0;
    // Start from row 1 if Excel has headers, or 0 if no headers (adjust as needed)
    // Assuming Excel format: Column A = ชื่อ-สกุล, Column B = ตำแหน่ง
    // And first row is header
    for (let i = 1; i < data.length; i++) {
      const name = data[i][0] ? data[i][0].toString().trim() : '';
      const position = data[i][1] ? data[i][1].toString().trim() : '';

      if (name && position) {
        if (!existingNames.includes(name)) {
          personnelSheet.appendRow([name, position]);
          existingNames.push(name); // Add to list to check against within this import
          addedCount++;
        } else {
          skippedCount++;
        }
      } else {
        skippedCount++; // Skip rows with missing name or position
      }
    }

    return { 
      success: true, 
      message: `นำเข้าข้อมูลสำเร็จ: เพิ่ม ${addedCount} รายการ, ข้าม ${skippedCount} รายการ (ซ้ำซ้อนหรือข้อมูลไม่ครบถ้วน)`,
      personnel: getPersonnelList()
    };

  } catch (e) {
    console.error("Error importing from Excel: " + e.toString() + "\nStack: " + e.stack);
    // Clean up temporary files if error occurs
    if (typeof spreadsheetId !== 'undefined' && spreadsheetId) {
        try { DriveApp.getFileById(spreadsheetId).setTrashed(true); } catch (err) {}
    }
    if (typeof tempSheetId !== 'undefined' && tempSheetId) {
        try { DriveApp.getFileById(tempSheetId).setTrashed(true); } catch (err) {}
    }
    return { success: false, message: "เกิดข้อผิดพลาดระหว่างการนำเข้า: " + e.message };
  }
}
