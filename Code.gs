// =========================================================
// ส่วนของระบบเข้ารหัส/ถอดรหัสข้อมูล
// =========================================================
const ENC_PREFIX = 'ENC:';
const ENC_KEY_PROP = 'ENC_KEY';
const DRIVE_FOLDER_NAME = 'ระบบยืมครุภัณฑ์_รูปบัตรประชาชน';

function getOrCreateEncryptionKey() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(ENC_KEY_PROP);
  if (!key) {
    key = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(ENC_KEY_PROP, key);
  }
  return key;
}

function generateKeyStream(seed, length) {
  const stream = [];
  let counter = 0;
  while (stream.length < length) {
    const block = Utilities.computeHmacSha256Signature(counter.toString(), seed);
    for (let i = 0; i < block.length && stream.length < length; i++) {
      stream.push(block[i] & 0xff);
    }
    counter++;
  }
  return stream;
}

function encryptText(plaintext) {
  if (plaintext === null || plaintext === undefined || plaintext === '') return '';
  try {
    const key = getOrCreateEncryptionKey();
    const bytes = Utilities.newBlob(String(plaintext)).getBytes();
    const nonce = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    const stream = generateKeyStream(key + ':' + nonce, bytes.length);
    const encrypted = [];
    for (let i = 0; i < bytes.length; i++) encrypted.push((bytes[i] & 0xff) ^ stream[i]);
    const nonceBytes = [];
    for (let i = 0; i < nonce.length; i++) nonceBytes.push(nonce.charCodeAt(i) & 0xff);
    const combined = nonceBytes.concat(encrypted);
    return ENC_PREFIX + Utilities.base64Encode(combined);
  } catch (e) {
    return 'ENC_ERR:' + e.toString();
  }
}

function decryptText(ciphertext) {
  if (ciphertext === null || ciphertext === undefined || ciphertext === '') return '';
  const str = String(ciphertext);
  if (!str.startsWith(ENC_PREFIX)) return str;
  try {
    const key = getOrCreateEncryptionKey();
    const combined = Utilities.base64Decode(str.substring(ENC_PREFIX.length));
    const nonceBytes = combined.slice(0, 16);
    const encrypted = combined.slice(16);
    let nonce = '';
    for (let i = 0; i < nonceBytes.length; i++) nonce += String.fromCharCode(nonceBytes[i] & 0xff);
    const stream = generateKeyStream(key + ':' + nonce, encrypted.length);
    const decrypted = [];
    for (let i = 0; i < encrypted.length; i++) decrypted.push((encrypted[i] & 0xff) ^ stream[i]);
    return Utilities.newBlob(decrypted).getDataAsString();
  } catch (e) {
    console.log('Decrypt error: ' + e.toString());
    return ''; 
  }
}

// =========================================================
// ส่วนของ Google Drive (สำหรับเก็บรูปบัตรประชาชน)
// =========================================================
function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function uploadImageToDrive(base64DataUrl, fileName) {
  try {
    if (!base64DataUrl || !base64DataUrl.startsWith('data:image/')) return '';
    const folder = getOrCreateFolder();
    const parts = base64DataUrl.split(',');
    const match = parts[0].match(/:(.*?);/);
    const contentType = match ? match[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getId();
  } catch (e) {
    console.log('Upload error: ' + e);
    return '';
  }
}

// =========================================================
// ส่วนของฟังก์ชันหลัก
// =========================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบยืมครุภัณฑ์ทางการแพทย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function formatDateThai(d) {
  if (!d) return '';
  if (d instanceof Date) {
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear() + 543;
    // แก้ไข: เลิกใช้ Backtick
    return day + '/' + month + '/' + year;
  }
  return d;
}

function saveRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ใบยืมครุภัณฑ์ทางการแพทย์');
    
    const itemsStr = data.items.map(i => i.name).join(', ');
    const equipStr = data.items.map(i => i.equipmentNumbers.join(', ')).join(' | ');
    const rawData = { items: data.items, address: data.address, phone: data.phone };
    
    const rowNum = sheet.getLastRow() + 1;
    // แก้ไข: เลิกใช้ Backtick
    const patientImgId = data.patientIdImage 
      ? uploadImageToDrive(data.patientIdImage, 'patient_' + rowNum + '_' + Date.now() + '.jpg') : '';
    const borrowerImgId = data.borrowerIdImage 
      ? uploadImageToDrive(data.borrowerIdImage, 'borrower_' + rowNum + '_' + Date.now() + '.jpg') : '';
    
    sheet.appendRow([
      sheet.getLastRow(),                              // A: ลำดับ
      new Date(),                                      // B: วันที่
      encryptText(data.patientName),                   // C
      encryptText(data.hn),                            // D
      encryptText(data.borrowerName),                  // E
      encryptText(data.relationship),                  // F
      encryptText(itemsStr),                           // G
      encryptText(equipStr),                           // H
      encryptText(data.address),                       // I
      encryptText(data.phone),                         // J
      data.totalDeposit,                               // K
      data.totalDeposit,                               // L
      '',                                              // M
      '',                                              // N
      encryptText(JSON.stringify(rawData)),            // O
      'ยังไม่ได้คืน',                                   // P
      patientImgId,                                    // Q: รูปบัตรประชาชนผู้ป่วย
      borrowerImgId                                    // R: รูปบัตรประชาชนผู้ยืม
    ]);
    
    return { success: true, message: 'บันทึกข้อมูลสำเร็จ (เข้ารหัสแล้ว)' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function getRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ใบยืมครุภัณฑ์ทางการแพทย์');
    const data = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[0] && !row[1] && !row[2]) continue;

      let items = [];
      let address = decryptText(row[8]) || '';
      let phone = decryptText(row[9]) || '';
      
      try {
        if (row[14]) {
          const decryptedJSON = decryptText(row[14]);
          if (decryptedJSON && !decryptedJSON.startsWith('ENC_ERR:')) {
            const parsed = JSON.parse(decryptedJSON);
            if (Array.isArray(parsed)) items = parsed;
            else {
              items = parsed.items || [];
              address = parsed.address || address;
              phone = parsed.phone || phone;
            }
          }
        }
      } catch (e) { console.log("Decrypt/Parse error row " + (i + 1)); }

      let itemsSummary = items.map(it => it.name + (it.qty > 1 ? ' x' + it.qty : '')).join(', ');
      
      records.push({
        id: row[0],
        date: formatDateThai(row[1]),
        patientName: decryptText(row[2]),
        hn: decryptText(row[3]),
        borrowerName: decryptText(row[4]),
        relationship: decryptText(row[5]),
        items: items,
        itemsSummary: itemsSummary,
        address: address,
        phone: phone,
        totalDeposit: row[10],
        returnDate: formatDateThai(row[12]),
        receiver: row[13] || '',
        status: row[15] || 'ยังไม่ได้คืน',
        patientIdImage: row[16] || '',   // Q
        borrowerIdImage: row[17] || ''   // R
      });
    }
    return records;
  } catch (error) {
    return { error: error.toString() };
  }
}

function updateReturnInfo(rowId, returnDate, receiver) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ใบยืมครุภัณฑ์ทางการแพทย์');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == rowId) {
        const rowNum = i + 1;
        let dateValue = '';
        if (returnDate) {
          const parts = returnDate.split('-');
          if (parts.length === 3) {
            dateValue = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          }
        }
        sheet.getRange(rowNum, 13).setValue(dateValue);
        sheet.getRange(rowNum, 14).setValue(receiver);
        sheet.getRange(rowNum, 16).setValue('คืนแล้ว');
        return { success: true, message: 'บันทึกข้อมูลการคืนสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการอัปเดต' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateRecordDetails(rowId, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ใบยืมครุภัณฑ์ทางการแพทย์');
    const dataRange = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == rowId) {
        const rowNum = i + 1;
        const itemsStr = data.items.map(it => it.name).join(', ');
        const equipStr = data.items.map(it => it.equipmentNumbers.join(', ')).join(' | ');
        const rawData = { items: data.items, address: data.address, phone: data.phone };
        
        sheet.getRange(rowNum, 3).setValue(encryptText(data.patientName));
        sheet.getRange(rowNum, 4).setValue(encryptText(data.hn));
        sheet.getRange(rowNum, 5).setValue(encryptText(data.borrowerName));
        sheet.getRange(rowNum, 6).setValue(encryptText(data.relationship));
        sheet.getRange(rowNum, 7).setValue(encryptText(itemsStr));
        sheet.getRange(rowNum, 8).setValue(encryptText(equipStr));
        sheet.getRange(rowNum, 9).setValue(encryptText(data.address));
        sheet.getRange(rowNum, 10).setValue(encryptText(data.phone));
        sheet.getRange(rowNum, 11).setValue(data.totalDeposit);
        sheet.getRange(rowNum, 12).setValue(data.totalDeposit);
        sheet.getRange(rowNum, 15).setValue(encryptText(JSON.stringify(rawData)));
        
        if (data.patientIdImage && data.patientIdImage.startsWith('data:image/')) {
          const imgId = uploadImageToDrive(data.patientIdImage, 'patient_' + rowNum + '_' + Date.now() + '.jpg');
          if (imgId) sheet.getRange(rowNum, 17).setValue(imgId);
        }
        if (data.borrowerIdImage && data.borrowerIdImage.startsWith('data:image/')) {
          const imgId = uploadImageToDrive(data.borrowerIdImage, 'borrower_' + rowNum + '_' + Date.now() + '.jpg');
          if (imgId) sheet.getRange(rowNum, 18).setValue(imgId);
        }
        
        return { success: true, message: 'บันทึกข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการแก้ไข' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
