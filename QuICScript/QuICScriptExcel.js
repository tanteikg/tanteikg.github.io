/**
 * @license
 * Author: pQCee
 * Description : QuICScript implementation in Office Add-ins for Excel
 *
 * Copyright pQCee 2024. All rights reserved
 *
 * “Commons Clause” License Condition v1.0
 *
 * The Software is provided to you by the Licensor under the License, as defined
 * below, subject to the following condition.
 *
 * Without limiting other conditions in the License, the grant of rights under
 * the License will not include, and the License does not grant to you, the
 * right to Sell the Software.
 *
 * For purposes of the foregoing, “Sell” means practicing any or all of the
 * rights granted to you under the License to provide to third parties, for a
 * fee or other consideration (including without limitation fees for hosting or
 * consulting/ support services related to the Software), a product or service
 * whose value derives, entirely or substantially, from the functionality of the
 * Software. Any license notice or attribution required by the License must also
 * include this Commons Clause License Condition notice.
 *
 * Software: QuICScript Office Add-in
 *
 * License: MIT, BSD-3-Clause
 *
 * Licensor: pQCee Pte Ltd
 */

// =======================================
// REGISTER EVENTS FOR HTML GUI COMPONENTS
// =======================================

Office.onReady((info) => {
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    // Setup actions 
    document.getElementById("btnCreateTable").onclick = handleCreateTable;
    document.getElementById("btnRunQuICScript").onclick = handleRunQuICScript;

  }
});

const M_TABLE_DEFAULT_DATA_ROWS = 10;

class QMLTemplate {
  constructor(intDataRows = M_TABLE_DEFAULT_DATA_ROWS) {
    // Initialise properties
    this.Colour = new Object();
    this.Worksheet = new Object();
    this.InstructionTable = new Object();
    this.ParamTable = new Object();
    this.MainTable = new Object();
    this.CommentsTable = new Object();

    // Initialise Colour
    this.#initialiseColour();

    // Calculate Template cell range locations
    this.#calculateWorksheet(Math.max(intDataRows, M_TABLE_DEFAULT_DATA_ROWS));
  }

  #initialiseColour() {
    this.Colour["Shade"] = new Object();

    // Cell Shade colours
    this.Colour.Shade["Grey"] = "#A5A5A5";
    this.Colour.Shade["Yellow"] = "#FFFF00";
  }

  /**
   * Function generates leftColumn, topRow, startCell, rightColumn, bottomRow,
   * endCell & range objects, then adds these to the target group object.
   *
   * @param {Object} objTargetGroup
   * Target group object to add generated objects.
   *
   * @param {string} charLeftColumn
   * Left column alphabet of target cell range.
   *
   * @param {number} intTopRow
   * Top row number of target cell range.
   *
   * @param {string} charRightColumn
   * Right column alphabet of target cell range.
   *
   * @param {number} intBottomRow
   * Bottom row number of target cell range.
   */
  #createCellRangeParams(
    objTargetGroup,
    charLeftColumn,
    intTopRow,
    charRightColumn,
    intBottomRow,
  ) {
    objTargetGroup["leftColumn"] = charLeftColumn;
    objTargetGroup["topRow"] = intTopRow;
    objTargetGroup["startCell"] = "".concat(charLeftColumn, intTopRow);
    objTargetGroup["rightColumn"] = charRightColumn;
    objTargetGroup["bottomRow"] = intBottomRow;
    objTargetGroup["endCell"] = "".concat(charRightColumn, intBottomRow);
    objTargetGroup["range"] = "".concat(
      objTargetGroup.startCell,
      ":",
      objTargetGroup.endCell,
    );
  }

  #calculateWorksheet(intDataRows) {
    // Audit Worksheet: Top-left cell (A1)
    this.Worksheet["leftColumn"] = "A";
    this.Worksheet["topRow"] = 1;
    this.Worksheet["startCell"] = "".concat(
      this.Worksheet.leftColumn,
      this.Worksheet.topRow,
    );

    // Audit Worksheet: right column
    this.Worksheet["rightColumn"] = "J";

    // Instruction Table
    this.#calculateInstructionTable();

    // Main Table
    this.#calculateMainTable(intDataRows);

    // Audit Table
    this.#calculateCommentsTable();

    // Message Params Table
    this.#calculateParamTable();

    // Audit Worksheet: Bottom-right cell
    this.Worksheet["bottomRow"] = this.CommentsTable.bottomRow;
    this.Worksheet["endCell"] = "".concat(
      this.Worksheet.startCell,
      ":",
      this.Worksheet.endCell,
    );
  }

  #calculateInstructionTable() {
    // Instruction Table: data and cell range
    this.InstructionTable["data"] = [
      ["Instructions:"],
      ["1. Auditor fills up Message Params and send workbook to client."],
      ["2. Client choose BTC/ETH in Crypto column."],
      ["3. Client fills up Wallet Address & Public Key."],
      ["4. Client sign Message and fills up Digital Signature."],
      ["5. Client sends workbook back to Auditor."],
      ["6. Auditor clicks Validate button to verify wallet ownership."],
    ];

    this.#createCellRangeParams(
      this.InstructionTable,
      this.Worksheet.leftColumn,
      this.Worksheet.topRow,
      this.Worksheet.leftColumn,
      this.InstructionTable.data.length,
    );
  }

  #calculateMainTable(intDataRows) {
    // Number of data rows in Main Table (excludes header row)
    this.MainTable["dataRows"] = intDataRows;

    // Main Table: cell range
    this.#createCellRangeParams(
      this.MainTable,
      this.Worksheet.leftColumn,
      this.InstructionTable.bottomRow + 2,
      this.Worksheet.rightColumn,
      this.InstructionTable.bottomRow + 2 + this.MainTable.dataRows,
    );

    // Main Table - Header: data, cell shades, and cell range
    this.MainTable["Header"] = new Object();

    this.MainTable.Header["data"] = [
      [
        "No.",
        "Crypto",
        "Wallet Address",
        "Public Key",
        "Message",
        "Digital Signature",
        "Valid Wallet",
        "Verified",
        "Blacklisted",
        "Balance",
      ],
    ];

    this.MainTable.Header["colours"] = [
      this.Colour.Shade.Grey,
      this.Colour.Shade.Yellow,
      this.Colour.Shade.Yellow,
      this.Colour.Shade.Yellow,
      this.Colour.Shade.Grey,
      this.Colour.Shade.Yellow,
      this.Colour.Shade.Grey,
      this.Colour.Shade.Grey,
      this.Colour.Shade.Grey,
      this.Colour.Shade.Grey,
    ];

    this.#createCellRangeParams(
      this.MainTable.Header,
      this.MainTable.leftColumn,
      this.MainTable.topRow,
      this.MainTable.rightColumn,
      this.MainTable.topRow,
    );

    // Main Table: Number of data columns
    this.MainTable["dataColumns"] = this.MainTable.Header.data[0].length;

    // Main Table - Data Section: cell range
    this.MainTable["Data"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Data,
      this.MainTable.leftColumn,
      this.MainTable.topRow + 1,
      this.MainTable.rightColumn,
      this.MainTable.bottomRow,
    );

    // Main Table - Column Index: zero-based index of columns
    this.MainTable["ColumnIndex"] = new Object();
    this.MainTable.ColumnIndex["SerialNumber"] =
      AuditTemplate.convertColToInt(this.MainTable.leftColumn) - 1;
    this.MainTable.ColumnIndex["Crypto"] =
      this.MainTable.ColumnIndex.SerialNumber + 1;
    this.MainTable.ColumnIndex["Address"] =
      this.MainTable.ColumnIndex.Crypto + 1;
    this.MainTable.ColumnIndex["PublicKey"] =
      this.MainTable.ColumnIndex.Address + 1;
    this.MainTable.ColumnIndex["Message"] =
      this.MainTable.ColumnIndex.PublicKey + 1;
    this.MainTable.ColumnIndex["Signature"] =
      this.MainTable.ColumnIndex.Message + 1;
    this.MainTable.ColumnIndex["ValidWallet"] =
      this.MainTable.ColumnIndex.Signature + 1;
    this.MainTable.ColumnIndex["ValidSignature"] =
      this.MainTable.ColumnIndex.ValidWallet + 1;
    this.MainTable.ColumnIndex["Blacklisted"] =
      this.MainTable.ColumnIndex.ValidSignature + 1;
    this.MainTable.ColumnIndex["Balance"] =
      this.MainTable.ColumnIndex.Blacklisted + 1;

    // Main Table - Column: alphabet of columns
    this.MainTable["Column"] = new Object();
    this.MainTable.Column["SerialNumber"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.SerialNumber + 1,
    );
    this.MainTable.Column["Crypto"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Crypto + 1,
    );
    this.MainTable.Column["Address"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Address + 1,
    );
    this.MainTable.Column["PublicKey"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.PublicKey + 1,
    );
    this.MainTable.Column["Message"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Message + 1,
    );
    this.MainTable.Column["Signature"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Signature + 1,
    );
    this.MainTable.Column["ValidWallet"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.ValidWallet + 1,
    );
    this.MainTable.Column["ValidSignature"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.ValidSignature + 1,
    );
    this.MainTable.Column["Blacklisted"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Blacklisted + 1,
    );
    this.MainTable.Column["Balance"] = AuditTemplate.convertIntToCol(
      this.MainTable.ColumnIndex.Balance + 1,
    );

    // Main Table - Serial Number Section: cell range
    this.MainTable["SerialNumber"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.SerialNumber,
      this.MainTable.Column.SerialNumber,
      this.MainTable.Data.topRow,
      this.MainTable.Column.SerialNumber,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Crypto Coin Section: cell range
    this.MainTable["Crypto"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Crypto,
      this.MainTable.Column.Crypto,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Crypto,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Crypto Address Section: cell range
    this.MainTable["Address"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Address,
      this.MainTable.Column.Address,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Address,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Crypto Public Key Section: cell range
    this.MainTable["PublicKey"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.PublicKey,
      this.MainTable.Column.PublicKey,
      this.MainTable.Data.topRow,
      this.MainTable.Column.PublicKey,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Message Section: cell range
    this.MainTable["Message"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Message,
      this.MainTable.Column.Message,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Message,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Signature Section: cell range
    this.MainTable["Signature"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Signature,
      this.MainTable.Column.Signature,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Signature,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Valid Wallet Section: cell range
    this.MainTable["ValidWallet"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.ValidWallet,
      this.MainTable.Column.ValidWallet,
      this.MainTable.Data.topRow,
      this.MainTable.Column.ValidWallet,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Valid Signature Section: cell range
    this.MainTable["ValidSignature"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.ValidSignature,
      this.MainTable.Column.ValidSignature,
      this.MainTable.Data.topRow,
      this.MainTable.Column.ValidSignature,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Blacklisted Section: cell range
    this.MainTable["Blacklisted"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Blacklisted,
      this.MainTable.Column.Blacklisted,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Blacklisted,
      this.MainTable.Data.bottomRow,
    );

    // Main Table - Account Balance Section: cell range
    this.MainTable["Balance"] = new Object();
    this.#createCellRangeParams(
      this.MainTable.Balance,
      this.MainTable.Column.Balance,
      this.MainTable.Data.topRow,
      this.MainTable.Column.Balance,
      this.MainTable.Data.bottomRow,
    );
  }

  #calculateParamTable() {
    // Message Params Table: data and cell range
    this.ParamTable["header"] = [["Message Params"]];

    this.ParamTable["description"] = [
      ["Seq. No."],
      ["Client Name"],
      ["Audit Date"],
    ];

    this.#createCellRangeParams(
      this.ParamTable,
      AuditTemplate.getColumnLeftOf(this.Worksheet.rightColumn),
      this.Worksheet.topRow + 1,
      this.Worksheet.rightColumn,
      this.Worksheet.topRow + 1 + this.ParamTable.description.length,
    );

    // Message Params Table - Header section: cell range
    this.ParamTable["Header"] = new Object();
    this.#createCellRangeParams(
      this.ParamTable.Header,
      this.ParamTable.leftColumn,
      this.ParamTable.topRow,
      this.ParamTable.rightColumn,
      this.ParamTable.topRow,
    );

    // Message Params Table - Description section: cell range
    this.ParamTable["Description"] = new Object();
    this.#createCellRangeParams(
      this.ParamTable.Description,
      this.ParamTable.leftColumn,
      this.ParamTable.topRow + 1,
      this.ParamTable.leftColumn,
      this.ParamTable.topRow + this.ParamTable.description.length,
    );

    // Message Params Table - Data section: cell range
    this.ParamTable["Data"] = new Object();
    this.#createCellRangeParams(
      this.ParamTable.Data,
      this.ParamTable.rightColumn,
      this.ParamTable.topRow + 1,
      this.ParamTable.rightColumn,
      this.ParamTable.topRow + this.ParamTable.description.length,
    );

    // Message Params Table - Data section: cell of SeqNo, ClientName, AuditDate
    this.ParamTable.Data["SeqNum"] = this.ParamTable.Data.startCell;
    this.ParamTable.Data["ClientName"] = "".concat(
      this.ParamTable.Data.leftColumn,
      this.ParamTable.Data.topRow + 1,
    );
    this.ParamTable.Data["AuditDate"] = "".concat(
      this.ParamTable.Data.leftColumn,
      this.ParamTable.Data.bottomRow,
    );
  }

  #calculateCommentsTable() {
    // Audit Comments Table has 10 rows
    this.CommentsTable["dataRows"] = 10;

    // Audit Comments Table: cell range
    this.#createCellRangeParams(
      this.CommentsTable,
      this.MainTable.leftColumn,
      this.MainTable.bottomRow + 2,
      this.MainTable.rightColumn,
      this.MainTable.bottomRow + 2 + this.CommentsTable.dataRows,
    );

    // Audit Comments Table: Header: cell range
    this.CommentsTable["Header"] = new Object();
    this.#createCellRangeParams(
      this.CommentsTable.Header,
      this.CommentsTable.leftColumn,
      this.CommentsTable.topRow,
      this.CommentsTable.rightColumn,
      this.CommentsTable.topRow,
    );

    // Audit Comments Table: Data: cell range
    this.CommentsTable["Data"] = new Object();
    this.#createCellRangeParams(
      this.CommentsTable.Data,
      this.CommentsTable.leftColumn,
      this.CommentsTable.topRow + 1,
      this.CommentsTable.rightColumn,
      this.CommentsTable.bottomRow,
    );
  }

  /**
   * Convert Excel column alphabet to column number.
   *
   * @param {string} charColAlpha - Single character containing column alphabet.
   * @returns {number} Number equivalent of column alphabet, where A = 1, B = 2, etc.
   */
  static convertColToInt(charColAlpha) {
    // This function does not receive arbitrary input from user.
    // Safe to assume the developer for this code will not pass in:
    // - zero-length string
    // - non-alphabet character
    // - double-alphabet string
    const ASCII_UPPER_CASE_A = "A".charCodeAt(0);
    return charColAlpha.toUpperCase().charCodeAt(0) - ASCII_UPPER_CASE_A + 1;
  }

  /**
   * Convert column number to Excel column alphabet.
   *
   * @param {number} intColNumber - Integer value of column number.
   * @returns {string} Character equivalent of column number, where 1 = A, 2 = B, etc.
   */
  static convertIntToCol(intColNumber) {
    // This function does not receive arbitrary input from user.
    // Similar to convertColToInt(), safe to assume developer does not pass in invalid values.
    const ASCII_UPPER_CASE_A = "A".charCodeAt(0);
    return String.fromCharCode(intColNumber + ASCII_UPPER_CASE_A - 1);
  }

  static getColumnLeftOf(charColAlpha) {
    return AuditTemplate.convertIntToCol(
      AuditTemplate.convertColToInt(charColAlpha) - 1,
    );
  }

  static getColumnRightOf(charColAlpha) {
    return AuditTemplate.convertIntToCol(
      AuditTemplate.convertColToInt(charColAlpha) + 1,
    );
  }
}

// =============
// BUTTON EVENTS
// =============

function handleCreateTable() {
  Excel.run((context) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    const intDataRows = M_TABLE_DEFAULT_DATA_ROWS; // Placeholder for user input
    generateAuditTableTemplate(selectedSheet, intDataRows);
    return context.sync();
  });
}

async function handleRunQuICScript() {
  await Excel.run(async (context) => {
    let targetRange = context.workbook.getSelectedRange();
    targetRange.load(["address","values"]);

    await context.sync();

    const QUICSTR = targetRange.values[0][0];
    const NUMQUBITS = targetRange.values[1][0];

    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    selectedSheet.getRange("A1:B1").values = [[QUICSTR,NUMQUBITS]];

    return context.sync();
  });
}

// ! Deprecated function. Code is kept here for reference purposes only.
// ! This code can be removed once we actively execute WASM components.
/*
// Test function to validate WebView2 support for Web Assembly execution
function handleTestWASM() {
  function wasmIsSupported() {
    try {
      if (
        typeof WebAssembly === "object" &&
        typeof WebAssembly.instantiate === "function"
      ) {
        const wasmBinary = Uint8Array.of(
          0x0,
          0x61,
          0x73,
          0x6d,
          0x01,
          0x00,
          0x00,
          0x00,
        );

        const module = new WebAssembly.Module(wasmBinary);

        if (module instanceof WebAssembly.Module) {
          return (
            new WebAssembly.Instance(module) instanceof WebAssembly.Instance
          );
        }
      }
    } catch (e) {
      console.error(e);
    }
    return false;
  }

  let testWASMMsg = document.getElementById("testWASMMsg");
  testWASMMsg.innerHTML = wasmIsSupported()
    ? "WebAssembly is supported"
    : "WebAssembly is not supported";
}
*/

// ================
// HELPER FUNCTIONS
// ================

/**
 * Converts a zero-indexed integer to an equivalent alphabet of the respective
 * column in Excel worksheet. Number more than 25 will be provided as a
 * multi-character column string. E.g., 0 -> "A", 26 -> "AA". The maximum number
 * of columns supported by Excel is 16,384 columns. The last column in Excel is
 * "XFD".
 *
 * @param {number} intColNumber - Zero-index integer column position.
 * @returns {string} - Returns empty string if parameter contains a number
 * beyond the range of 0 to 16,384; otherwise returns column name of Excel
 * worksheet.
 */
function intToExcelColumnAlpha(intColNumber) {
  const ALPHABETS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const BASE26 = 26;
  const EXCEL_MAX_COLUMNS = 16384;

  // Sanity checks
  if (intColNumber < 0 || intColNumber >= EXCEL_MAX_COLUMNS) {
    return "";
  } else {
    let columnName = "";

    do {
      columnName = "".concat(ALPHABETS[intColNumber % BASE26], columnName);
      intColNumber = Math.floor(intColNumber / BASE26) - 1;
    } while (intColNumber >= 0);

    return columnName;
  }
}

function generateAuditTableTemplate(objWS, intDataRows) {
  /**
   * Converts value for Excel row height or column width from pixel to font
   * point. The conversion is achieved by using an approximation, where font
   * point = pixel * 0.75.
   *
   * @inner
   * @param {number} intPixelSize - Value for Excel row height or column width in pixels.
   * @returns {number} Equivalent width (floating point) value in Excel font points.
   */
  function pixelToPoint(intPixelSize) {
    return intPixelSize * 0.75;
  }

  //
  // Calculate cell range of Main Table
  //
  const AT = Object.freeze(new AuditTemplate(intDataRows));

  //
  // Populate Instructions Table
  //
  objWS.getRange(AT.InstructionTable.range).values = AT.InstructionTable.data;

  // =============================
  // MAIN TABLE BELOW INSTRUCTIONS
  // =============================
  // MAIN TABLE: HEADER
  objWS.getRange(AT.MainTable.Header.range).values = AT.MainTable.Header.data;
  objWS.getRange(AT.MainTable.Header.range).format.font.bold = true;

  // TODO: optimise multiple format calls into one set()
  for (
    let col = AT.MainTable.ColumnIndex.SerialNumber, row = AT.MainTable.topRow;
    col <= AT.MainTable.ColumnIndex.Balance;
    col++
  ) {
    objWS.getRange("".concat(convertIntToCol(col + 1), row)).format.fill.color =
      AT.MainTable.Header.colours[col];
  }

  // MAIN TABLE
  addBorderLines(objWS.getRange(AT.MainTable.range));
  objWS.getRange(AT.MainTable.range).format.horizontalAlignment = "Center";
  objWS.getRange(AT.MainTable.Header.range).numberFormat = "@";

  // MAIN TABLE: DATA
  objWS.getRange(AT.MainTable.SerialNumber.range).numberFormat = "0";
  objWS.getRange(
    "".concat(AT.MainTable.Crypto.startCell, ":", AT.MainTable.Balance.endCell),
  ).numberFormat = "@";

  // MAIN TABLE: DATA apply word-wrap from columns C to F
  objWS.getRange(
    "".concat(
      AT.MainTable.Address.startCell,
      ":",
      AT.MainTable.Signature.endCell,
    ),
  ).format.wrapText = true;

  // MAIN TABLE: DATA left 2nd Column (Crypto)
  objWS.getRange(AT.MainTable.Crypto.range).dataValidation.clear();
  objWS.getRange(AT.MainTable.Crypto.range).dataValidation.rule = {
    list: { inCellDropDown: true, source: "BTC,ETH" },
  };

  // MAIN TABLE: VALIDATION COLUMNS (Valid Wallet + Verified Signature)
  const M_TABLE_VAL_RANGE = "".concat(
    AT.MainTable.ValidWallet.startCell,
    ":",
    AT.MainTable.ValidSignature.endCell,
  );
  objWS.getRange(M_TABLE_VAL_RANGE).conditionalFormats.clearAll();
  const trueConditionalFormat = objWS
    .getRange(M_TABLE_VAL_RANGE)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  trueConditionalFormat.textComparison.format.font.color = "#006100";
  trueConditionalFormat.textComparison.format.fill.color = "#C6EFCE";
  trueConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "Yes",
  };
  const falseConditionalFormat = objWS
    .getRange(M_TABLE_VAL_RANGE)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  falseConditionalFormat.textComparison.format.font.color = "#9C0006";
  falseConditionalFormat.textComparison.format.fill.color = "#FFC7CE";
  falseConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "No",
  };

  // MAIN TABLE: VALIDATION COLUMN (Blacklisted)
  objWS.getRange(AT.MainTable.Blacklisted.range).conditionalFormats.clearAll();
  const noBLConditionalFormat = objWS
    .getRange(AT.MainTable.Blacklisted.range)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  noBLConditionalFormat.textComparison.format.font.color = "#006100";
  noBLConditionalFormat.textComparison.format.fill.color = "#C6EFCE";
  noBLConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "No",
  };
  const yesBLConditionalFormat = objWS
    .getRange(AT.MainTable.Blacklisted.range)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  yesBLConditionalFormat.textComparison.format.font.color = "#9C0006";
  yesBLConditionalFormat.textComparison.format.fill.color = "#FFC7CE";
  yesBLConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "Yes",
  };

  // MAIN TABLE: Fill index column
  let indexNumbers = new Array(intDataRows);
  for (let i = 0; i < intDataRows; i++) {
    indexNumbers[i] = [i + 1];
  }
  objWS.getRange(AT.MainTable.SerialNumber.range).values = indexNumbers;

  // ==============================================
  // MESSAGE PARAMS TABLE AT TOP-RIGHT OF WORKSHEET
  // ==============================================
  // TODO: optimise multiple format calls into one set()
  objWS.getRange(AT.ParamTable.startCell).values = AT.ParamTable.header;
  objWS.getRange(AT.ParamTable.Header.range).merge(false);
  objWS.getRange(AT.ParamTable.Header.range).format.fill.color =
    AT.Colour.Shade.Yellow;
  objWS.getRange(AT.ParamTable.Header.range).format.font.bold = true;
  objWS.getRange(AT.ParamTable.Header.range).format.horizontalAlignment =
    "Center";
  objWS.getRange(AT.ParamTable.Description.range).values =
    AT.ParamTable.description;
  objWS.getRange(AT.ParamTable.Data.range).numberFormat = "@";
  objWS.getRange(AT.ParamTable.Data.SeqNum).values = [
    [Math.floor(Math.random() * 9000) + 1000],
  ];
  objWS.getRange(AT.ParamTable.Data.ClientName).values = [["Company A"]];
  objWS.getRange(AT.ParamTable.Data.AuditDate).values = [[todayDate()]];
  addBorderLines(objWS.getRange(AT.ParamTable.range));

  // =====================================
  // AUDIT COMMENTS TABLE BELOW MAIN TABLE
  // =====================================
  // AUDIT COMMENTS TABLE: HEADER
  objWS.getRange(AT.CommentsTable.Header.range).merge(false);
  addBorderLines(objWS.getRange(AT.CommentsTable.Header.range));
  objWS.getRange(AT.CommentsTable.Header.range).set({
    format: {
      horizontalAlignment: "Left",
      font: {
        bold: true,
      },
      fill: {
        color: AT.Colour.Shade.Yellow,
      },
    },
  });
  objWS.getRange(AT.CommentsTable.Header.startCell).values = [
    ["Audit Comments"],
  ];

  // AUDIT COMMENTS TABLE: DATA
  objWS.getRange(AT.CommentsTable.Data.range).merge(false);
  addBorderLines(objWS.getRange(AT.CommentsTable.Data.range));
  objWS.getRange(AT.CommentsTable.Data.range).format.horizontalAlignment =
    "Left";

  // ================================
  // WORKSHEET RANGE FORMAT SETTINGS
  // ===============================
  objWS.getRange(AT.Worksheet.range).set({
    format: {
      font: {
        color: "#000000",
        name: "Calibri",
        size: 10,
      },
      verticalAlignment: "Center",
    },
  });

  // Only Audit Comments Table: DATA need to be Top-justified
  objWS.getRange(AT.CommentsTable.Data.range).format.verticalAlignment = "Top";

  // =======================================
  // WORKSHEET COLUMN WIDTHS AND ROW HEIGHTS
  // =======================================
  // TODO: optimise multiple format calls into one set()
  objWS.getRange("A1").format.columnWidth = pixelToPoint(29);
  objWS.getRange("B1").format.columnWidth = pixelToPoint(44);
  objWS.getRange("C1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("D1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("E1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("F1").format.columnWidth = pixelToPoint(265);
  objWS.getRange("G1").format.columnWidth = pixelToPoint(74);
  objWS.getRange("H1").format.columnWidth = pixelToPoint(74);
  objWS.getRange("I1").format.columnWidth = pixelToPoint(74);
  objWS.getRange("J1").format.columnWidth = pixelToPoint(78);
  // Note: If you manually set the rowHeight, Excel no longer autofits rows
  //       to contents of cells with "wrapText = true". The way to do this
  //       is to not set the rowHeight programmatically.
  // objWS.getRange(WS_RANGE).format.rowHeight = pixelToPoint(17);
}

