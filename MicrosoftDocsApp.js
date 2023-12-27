/**
 * GitHub  https://github.com/tanaikech/MicrosoftDocsApp<br>
 * Library name
 * @type {string}
 * @const {string}
 * @readonly
 */
var appName = "MicrosoftDocsApp";

/**
 * ### Description
 * Set file ID.
 * 
 * @param {String} fileId
 * @return {MicrosoftDocsApp}
 */
function setFileId(fileId) {
  this.obj = new MicrosoftDocsApp_(fileId);
  return this;
}

/**
* ### Description
* Get Class Document.
*
* @returns {DocumentApp.Document} document
*/
function getDocument() {
  return this.obj.getDocument();
}

/**
* ### Description
* Get Class Spreadsheet.
*
* @returns {SpreadsheetApp.Spreadsheet} spreadsheet
*/
function getSpreadsheet() {
  return this.obj.getSpreadsheet();
}

/**
* ### Description
* Get Class Presentation.
*
* @returns {SlidesApp.Presentation} presentation
*/
function getSlide() {
  return this.obj.getSlide();
}

/**
* ### Description
* Save XLSX data.
*
* @return {void}
*/
function saveAndClose() {
  return this.obj.saveAndClose();
}

/**
* ### Description
* Delete temporal file.
*
* @return {void}
*/
function end() {
  return this.obj.end();
}


/**
 * Return Class Document, Class Spreadsheet, or Class Presentation of Google Doc files from the file ID of Microsoft Doc files of Word, Excel, or PowerPoint.
 */
class MicrosoftDocsApp_ {

  /**
  * @param {String} fileId
  */
  constructor(fileId) {
    this.fileId = fileId;
    this.mimeTypesOfGoogleDocs = [MimeType.GOOGLE_DOCS, MimeType.GOOGLE_SHEETS, MimeType.GOOGLE_SLIDES];
    this.mimeTypesOfMicroSoftDocs = [MimeType.MICROSOFT_WORD, MimeType.MICROSOFT_EXCEL, MimeType.MICROSOFT_POWERPOINT];
    this.m2gObj = this.mimeTypesOfMicroSoftDocs.reduce((o, e, i) => (o[e] = this.mimeTypesOfGoogleDocs[i], o), {});
    this.g2mObj = this.mimeTypesOfGoogleDocs.reduce((o, e, i) => (o[e] = this.mimeTypesOfMicroSoftDocs[i], o), {});
    this.srcMimeType = DriveApp.getFileById(this.fileId).getMimeType();
    this.tempMimeType = "";
    this.convert = this.mimeTypesOfMicroSoftDocs.includes(this.srcMimeType);
    this.headers = { authorization: "Bearer " + ScriptApp.getOAuthToken() };
  }

  /**
  * ### Description
  * Get Class Document.
  *
  * @returns {DocumentApp.Document} document
  */
  getDocument() {
    return this.getDoc_();
  }

  /**
  * ### Description
  * Get Class Spreadsheet.
  *
  * @returns {SpreadsheetApp.Spreadsheet} spreadsheet
  */
  getSpreadsheet() {
    return this.getDoc_();
  }

  /**
  * ### Description
  * Get Class Presentation.
  *
  * @returns {SlidesApp.Presentation} presentation
  */
  getSlide() {
    return this.getDoc_();
  }

  /**
  * ### Description
  * Save XLSX data.
  *
  * @return {void}
  */
  saveAndClose() {
    if (!this.convert) return;
    if ([MimeType.MICROSOFT_WORD, MimeType.MICROSOFT_POWERPOINT].includes(this.srcMimeType)) {
      this.doc.saveAndClose();
    } else {
      SpreadsheetApp.flush();
    }
    Drive.Files.update({}, this.fileId, UrlFetchApp.fetch(this.convEndpoint, { headers: this.headers }).getBlob(), { supportsAllDrives: true });
  }

  /**
  * ### Description
  * Delete temporal file.
  *
  * @return {void}
  */
  end() {
    if (!this.convert) return;
    Utilities.sleep(1000);
    Drive.Files.remove(this.tempId, { supportsAllDrives: true });
  }

  /**
  * @returns {SpreadsheetApp.Spreadsheet|DocumentApp.Document|SlidesApp.Presentation} doc
  */
  getDoc_() {
    let id;
    if (this.convert) {
      const temp = Drive.Files.copy({ name: "temp", mimeType: this.m2gObj[this.srcMimeType] }, this.fileId, { supportsAllDrives: true, fields: "id,exportLinks,mimeType" });
      this.convEndpoint = temp.exportLinks[this.g2mObj[temp.mimeType]];
      this.tempId = temp.id;
      this.tempMimeType = temp.mimeType;
      id = temp.id;
    } else {
      if (![...this.mimeTypesOfGoogleDocs, ...this.mimeTypesOfMicroSoftDocs].includes(this.srcMimeType)) {
        throw new Error(`This file ID "${this.fileId}" is not Microsoft Docs file or Google Docs file.`);
      }
      id = this.fileId;
    }
    if ([this.tempMimeType, this.srcMimeType].includes(MimeType.GOOGLE_DOCS)) {
      this.doc = DocumentApp.openById(id);
    } else if ([this.tempMimeType, this.srcMimeType].includes(MimeType.GOOGLE_SHEETS)) {
      this.doc = SpreadsheetApp.openById(id);
    } else if ([this.tempMimeType, this.srcMimeType].includes(MimeType.GOOGLE_SLIDES)) {
      this.doc = SlidesApp.openById(id);
    }
    return this.doc;
  }
}
