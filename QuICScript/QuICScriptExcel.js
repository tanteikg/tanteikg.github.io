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

// QuICScript functions
//

var inited = 0;
var message = "";

async function resetQuICScript()
{
	Module._QuICScript_end();
	inited = 0;
	message = "state is cleared";
}

async function runQuICScript(Qcir, Qnum) 
{
	if (!inited)
	{
		Module._QuICScript_begin(Qnum);
		inited = Qnum;
		message = "State is reset, working on "+Qnum+" Qubits\n";
	}
	else
	{
		if (inited != Qnum)
		{
			Module._QuICScript_end();
			Module._QuICScript_begin(Qnum);
			inited = Qnum;
			message = "State is reset, working on "+Qnum+" Qubits\n";
		}
	}

	resultstate= Module.ccall('QuICScript_cont','string',['number','string','number','number','number','number','number','number','number','number'],[Qnum,Qcir,1,0,0,0,0,0,1,0]);
	message = resultstate + "---\n" + message;  

} 


// =============
// BUTTON EVENTS
// =============

function handleCreateTable() {
	Excel.run((context) => {
		let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

		selectedSheet.getRange("A1:B1").values = [["hello","world"],];
		return context.sync();
	});
}

async function handleRunQuICScript() {
	await Excel.run(async (context) => {
		let targetRange = context.workbook.getSelectedRange();
		targetRange.load(["address","values"]);

		await context.sync();

		const QUICSTR = targetRange.values[0][0];
		const NUMQUBITS = targetRange.values[0][1];

		runQuICScript(QUICSTR,NUMQUBITS);

		let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
		selectedSheet.getRange("A1:B1").values = [[QUICSTR,NUMQUBITS]];
		selectedSheet.getRange("A2:A2").values = [[message]];

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

