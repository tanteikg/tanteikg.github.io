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
var result = "";

async function resetQuICScript()
{
	Module._QuICScript_end();
	inited = 0;
	message = "state is cleared";
	result = "";
}

async function runQuICScript(Qcir, Qnum, theta, phi, lamda) 
{
	if (!inited)
	{
		Module._QuICScript_begin(Qnum);
		inited = Qnum;
		message = "State is reset, working on "+Qnum+" Qubits\n";
		result = "";
	}
	else
	{
		if (inited != Qnum)
		{
			Module._QuICScript_end();
			Module._QuICScript_begin(Qnum);
			inited = Qnum;
			message = "State is reset, working on "+Qnum+" Qubits\n";
			result = "";
		}
	}

	resultstate= Module.ccall('QuICScript_cont','string',['number','string','number','number','number','number','number','number','number','number'],[Qnum,Qcir,theta,0,phi,0,lamda,0,1,0]);
	message = Qcir + " is run";
	result = resultstate;

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

		resetQuICScript();
		return context.sync();
	}
}
async function handleRunQuICScript() {
	await Excel.run(async (context) => {
		let targetRange = context.workbook.getSelectedRange();
		targetRange.load(["address","values"]);

		await context.sync();
		const NUMQUBITS = targetRange.values[0][0];
		var QUICSTR;
		var line = 1;
		var theta;
		var phi;
		var lamda;

		let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
		let resultbox = document.getElementById("resultbox").value;
		selectedSheet.getRange(resultbox).values = [[""]];
		while ((QUICSTR = targetRange.values[line][0]) != null)
		{
			if ((theta = targetRange.values[line][1]) == null)
				theta = 0;
			if ((phi = targetRange.values[line][2]) == null)
				phi = 0;
			if ((lamda = targetRange.values[line][3]) == null)
				lamda = 0;

			runQuICScript(QUICSTR,NUMQUBITS,theta,phi,lamda);
			document.getElementById("message").innerText = "Line"+line+": " + message;
			selectedSheet.getRange(resultbox).values = [[result]];
			line++;
		}


		return context.sync();
	});
}


