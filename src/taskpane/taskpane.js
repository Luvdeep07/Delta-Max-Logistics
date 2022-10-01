/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { mainModule } = require("process");

/* global console, document, Excel, Office */

var selectedRangeSort;
var selectedRangeSums = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("start").onclick = start;
    document.getElementById("sort-refresh").onclick = sortRefresh;
    document.getElementById("select-to-sum").onclick = selectToSum;
    document.getElementById("restart").onclick = restart;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
  }
});

async function restart() {
  await Excel.run(async (context) => {
    document.getElementById("selected-field-to-sort").style.display = "none";
    document.getElementById("selected").style.display = "none";
    document.getElementById("dot").style.display = "none";

    document.getElementById("select-to-sum").style.display = "none";
    document.getElementById("sort-refresh").style.display = "none";

    document.getElementById("selected-field-to-sum").style.display = "none";
    document.getElementById("fields").style.display = "none";

    document.getElementById("start").style.display = "initial";
    document.getElementById("restart").style.display = "none";

    document.getElementById("fields").textContent = "...";
    document.getElementById("selected").textContent = "Not selected";

    selectedRangeSums = [];

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof Office.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function selectToSum() {
  await Excel.run(async (context) => {
    var selectedRangeSum = context.workbook.getSelectedRange();
    selectedRangeSum.load("address");
    selectedRangeSum.load("text");
    await context.sync();
    var address = selectedRangeSum.address.split("!")[1];
    var col = address.match(/[a-zA-Z]/);
    selectedRangeSums.push([selectedRangeSum.text[0][0], col]);

    var val = "";
    for (var i = 0; i < selectedRangeSums.length; i++) {
      let sumOnVal = selectedRangeSums[i][0];
      val += sumOnVal;

      if (i < selectedRangeSums.length - 1) val += ", ";
    }

    document.getElementById("fields").textContent = val;

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof Office.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function start() {
  await Excel.run(async (context) => {
    document.getElementById("start").style.display = "none";

    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("address");
    selectedRange.load("text");
    await context.sync();
    let sortOnVal = selectedRange.text[0][0];
    let sortOnAddress = selectedRange.address.split("!")[1];
    selectedRangeSort = [sortOnVal, sortOnAddress];

    document.getElementById("selected-field-to-sort").style.display = "initial";
    document.getElementById("selected").textContent = sortOnVal;
    document.getElementById("selected").style.display = "initial";
    document.getElementById("dot").style.display = "initial";

    document.getElementById("select-to-sum").style.display = "initial";
    document.getElementById("sort-refresh").style.display = "initial";

    document.getElementById("selected-field-to-sum").style.display = "initial";
    document.getElementById("fields").style.display = "initial";

    document.getElementById("restart").style.display = "initial";

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof Office.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function sortRefresh() {
  await Excel.run(async (context) => {
    var worksheets = context.workbook.worksheets;
    const currentWorksheet = worksheets.getActiveWorksheet();

    let range = currentWorksheet.getUsedRange();
    range.load("text");
    await context.sync();
    let rangeValues = range.text;

    let sortOn = selectedRangeSort[1];
    let sortOnRow = sortOn.match(/\d+/);
    let sortOnCol = sortOn.match(/[a-zA-Z]/);
    let sortOnVal = selectedRangeSort[0];

    let storeSorted = {};
    var afterSortOn = false;
    let header = [[]];
    for (var i in rangeValues) {
      let row = rangeValues[i];

      if (!afterSortOn) {
        afterSortOn = sortOnVal == row[colA1ToIndex(sortOnCol) - 1];
        if (afterSortOn) {
          for (var k in row) {
            header[0].push(row[k]);
          }
        }
        continue;
      }

      let key = (row[colA1ToIndex(sortOnCol) - 1] + "").trim();
      if (key == "") key = "EMPTY";
      if (key.length > 31) key = key.slice(0, 31);
      else key = key[0].toUpperCase() + key.slice(1);

      key = key.split(":").join("");
      key = key.split("\\").join("");
      key = key.split("/").join("");
      key = key.split("?").join("");
      key = key.split("*").join("");
      key = key.split("[").join("");
      key = key.split("]").join("");

      if (!(key in storeSorted)) {
        storeSorted[key] = [];
      }

      let arr = [];
      for (var j in row) {
        arr.push(row[j]);
      }
      storeSorted[key].push(arr);
    }

    for (const [key, val] of Object.entries(storeSorted)) {
      let wb = context.workbook.worksheets.getItemOrNullObject(key);
      wb.load("isNullObject");
      await context.sync();

      if (wb.isNullObject) {
        worksheets.add(key);
      }

      var ws = worksheets.getItem(key);
      ws.getRange().clear();

      let headerRange = "A1:" + colName(header[0].length - 1) + "1";
      let hRange = ws.getRange(headerRange);
      hRange.values = header;
      hRange.format.fill.color = "#244062";
      hRange.format.font.color = "white";
      hRange.format.font.bold = true;
      hRange.format.font.size = 12;

      let insertRange = "A2:" + colName(val[0].length - 1) + (val.length + 1);
      let wbRange = ws.getRange(insertRange);
      wbRange.values = val;
      wbRange.format.font.size = 12;
      var sums = [];
      for (var x = 0; x < selectedRangeSums.length; x++) {
        var colA1 = selectedRangeSums[x][1];
        var col = colA1ToIndex(colA1);
        sums.push(0);

        for (var row in val) {
          if (/^\d+$/.test(val[row][col - 1])) sums[x] += Number(val[row][col - 1]);
        }

        let sumRange = colA1 + (val.length + 2);
        let sRange = ws.getRange(sumRange);
        sRange.values = [[sums[x]]];
        sRange.format.font.bold = true;
        sRange.format.font.size = 12;
      }

      wbRange.format.autofitColumns();
      wbRange.format.autofitRows();
    }

    document.getElementById("selected-field-to-sort").style.display = "none";
    document.getElementById("selected").style.display = "none";
    document.getElementById("dot").style.display = "none";

    document.getElementById("select-to-sum").style.display = "none";
    document.getElementById("sort-refresh").style.display = "none";

    document.getElementById("selected-field-to-sum").style.display = "none";
    document.getElementById("fields").style.display = "none";

    document.getElementById("start").style.display = "initial";
    document.getElementById("restart").style.display = "none";

    document.getElementById("fields").textContent = "...";
    document.getElementById("selected").textContent = "Not selected";

    selectedRangeSums = [];

    // const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    // const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    // expensesTable.name = "ExpensesTable";

    // expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    // expensesTable.rows.add(null /*add at the end*/, [
    //   ["1/1/2017", "The Phone Company", "Communications", "120"],
    //   ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
    //   ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
    //   ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
    //   ["1/11/2017", "Bellows College", "Education", "350.1"],
    //   ["1/15/2017", "Trey Research", "Other", "135"],
    //   ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    // ]);

    // expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    // expensesTable.getRange().format.autofitColumns();
    // expensesTable.getRange().format.autofitRows();

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof Office.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

function colName(n) {
  n += 1;
  n -= 1;
  var ordA = "A".charCodeAt(0);
  var ordZ = "Z".charCodeAt(0);
  var len = ordZ - ordA + 1;

  var s = "";
  while (n >= 0) {
    s = String.fromCharCode((n % len) + ordA) + s;
    n = Math.floor(n / len) - 1;
  }

  return s;
}

function colA1ToIndex(colA1) {
  let result = 0;

  let strLen = colA1.length;

  for (let i = 0; i < strLen; i++) {
    result += (colA1[i].charCodeAt() - 64) * Math.pow(26, strLen - i - 1);
  }

  return result;
}
