

var CONFIG = {
  rowsPerDay: 5, 
  year: 2025, 
  currency: "€", 
  colors: {
    header: "#37474F", week: "#546E7A", 
    dayBlock: "#FFFFFF", dayBorder: "#000000",
    subHeader: "#CFD8DC",
    general: "#E3F2FD", wants: "#FCE4EC", culture: "#E8F5E9", extra: "#FFF3E0",
    warning: "#FFEBEE" 
  },
  // --- USER CONFIGURATION START ---
  incomeItems: ["Salary A", "Salary B", "Variable / Extra"],
  fixedExpenses: ["Mortgage", "Home Insurance", "Phone/Internet", "Condo Fees", "School", "Netflix", "Other"],
 
  categories: [
    "Needs(Groceries)", "Needs(Pharmacy)", "Needs(Transport)", "Needs(Kids)", "Needs(Others)", 
    "Optional(Coffee/Bars)", "Optional(Dining Out)", "Optional(Beauty)", "Optional(Shopping)", 
    "Culture(Books)", "Culture(Music & Movies)", "Culture(Shows/Events)",  
    "Extra(Gifts)", "Extra(Household)", "Extra(Other)"
  ],
  // --- USER CONFIGURATION END ---
  
  months: ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
};

// --- DYNAMIC LAYOUT CALCULATIONS ---
var LAYOUT = {
  incomeStart: 5,
  incomeEnd: 5 + CONFIG.incomeItems.length - 1, 
  fixedStart: 5,
  fixedEnd: 5 + CONFIG.fixedExpenses.length - 1,
  incomeTotalRow: 5 + CONFIG.incomeItems.length,
  fixedTotalRow: 5 + CONFIG.fixedExpenses.length,
  statusRow: Math.max(5 + CONFIG.incomeItems.length, 5 + CONFIG.fixedExpenses.length) + 3,
  logStartRow: Math.max(5 + CONFIG.incomeItems.length, 5 + CONFIG.fixedExpenses.length) + 10 
};

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Kakeibo Settings", [
    {name: "🔄 Update Configuration (Keep Data)", functionName: "updateKakeiboSafe"},
    {name: "⚠️ Factory Reset (Wipe All)", functionName: "setupFixedBlockKakeibo"}
  ]);
  var m = CONFIG.months;
  var sheet = ss.getSheetByName(m[new Date().getMonth()]);
  if (sheet) sheet.activate();
}

function setupFixedBlockKakeibo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.setSpreadsheetLocale('it_IT'); 
  
  var response = ui.prompt('FACTORY RESET', 'Enter Fiscal Year (e.g., 2025). WARNING: This deletes all data.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  CONFIG.year = parseInt(response.getResponseText());

  buildSystem(ss, false);
}

function updateKakeiboSafe() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.setSpreadsheetLocale('it_IT'); 

  var jan = ss.getSheetByName("JAN");
  if(jan) {
    var title = jan.getRange("B2").getValue(); 
    var parts = title.split(" ");
    if(parts.length > 0) {
      var detectedYear = parseInt(parts[parts.length-1]);
      if(!isNaN(detectedYear)) CONFIG.year = detectedYear;
    }
  }

  var response = ui.alert('UPDATE CONFIG', `Update lists for Year ${CONFIG.year}?\nThis will rebuild the layout but preserve your entered values.`, ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) return;

  buildSystem(ss, true);
}

function buildSystem(ss, preserveData) {
  CONFIG.months.forEach(function(month, index) {
    var sheet = ss.getSheetByName(month);
    var savedData = null;
    if (preserveData && sheet) { savedData = scrapeSheetData(sheet); }
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(month, index + 1); 
    createBlockMonth(sheet, month, index);
    if (preserveData && savedData) { restoreSheetData(sheet, savedData); }
  });

  createDashboard(ss); // Build Last to avoid REF errors

  var sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) ss.deleteSheet(sheet1);
  protectSystem(ss);
  
  ss.toast(preserveData ? "✅ Configuration Updated! Data Preserved." : "✅ Kakeibo Reset Successful.");
}

function scrapeSheetData(sheet) {
  var data = { incomes: {}, fixed: {}, savingsGoal: 0, logs: [] };
  try {
    data.savingsGoal = sheet.getRange("I5").getValue();
    for(var r=4; r<15; r++) {
       var labelInc = sheet.getRange(r, 2).getValue();
       var valInc = sheet.getRange(r, 3).getValue();
       if(labelInc && typeof valInc === 'number') data.incomes[labelInc] = valInc;
       var labelFix = sheet.getRange(r, 5).getValue();
       var valFix = sheet.getRange(r, 6).getValue();
       if(labelFix && typeof valFix === 'number') data.fixed[labelFix] = valFix;
    }
    var lastRow = sheet.getLastRow();
    for(var r=15; r<=lastRow; r++) {
       var rowVals = sheet.getRange(r, 2, 1, 7).getValues()[0]; 
       var amt = rowVals[5];
       if(amt !== "" && amt !== 0 && amt !== null) {
         data.logs.push({ date: rowVals[0], who: rowVals[2], item: rowVals[3], cat: rowVals[4], amt: amt, ref: rowVals[6] });
       }
    }
  } catch(e) { Logger.log("Error scraping " + sheet.getName() + ": " + e); }
  return data;
}

function restoreSheetData(sheet, data) {
  CONFIG.incomeItems.forEach(function(item, i) {
    if(data.incomes[item]) sheet.getRange(LAYOUT.incomeStart + i, 3).setValue(data.incomes[item]);
  });
  CONFIG.fixedExpenses.forEach(function(item, i) {
    if(data.fixed[item]) sheet.getRange(LAYOUT.fixedStart + i, 6).setValue(data.fixed[item]);
  });
  sheet.getRange("I5").setValue(data.savingsGoal);

  var logData = sheet.getRange(LAYOUT.logStartRow + 1, 2, 365 * CONFIG.rowsPerDay, 7).getValues();
  data.logs.forEach(function(entry) {
    if(!entry.date) return;
    var entryDateStr = new Date(entry.date).toDateString();
    for(var i=0; i<logData.length; i++) {
       var rowDate = logData[i][0];
       if(rowDate instanceof Date && rowDate.toDateString() === entryDateStr) {
         if(logData[i][5] === "") { 
           logData[i][2] = entry.who; logData[i][3] = entry.item; logData[i][4] = entry.cat; logData[i][5] = entry.amt; logData[i][6] = entry.ref;
           break; 
         }
       }
    }
  });
  sheet.getRange(LAYOUT.logStartRow + 1, 2, logData.length, 7).setValues(logData);
}

// --- STANDARD BUILDERS ---

function createBlockMonth(sheet, monthName, monthIndex) {
  sheet.clear(); sheet.setHiddenGridlines(true);
  formatHeader(sheet.getRange("B2:L2"), `KAKEIBO: ${monthName} ${CONFIG.year}`, CONFIG.colors.header);
  setupTopSection(sheet); 
  var startRow = LAYOUT.logStartRow;
  sheet.getRange(startRow, 2, 1, 7).setValues([["Date", "Day", "Who?", "Item / Description", "Category", "Amount", "Reflection"]])
       .setFontWeight("bold").setBackground(CONFIG.colors.header).setFontColor("white");

  var currentRow = startRow + 1;
  var daysInMonth = new Date(CONFIG.year, monthIndex + 1, 0).getDate();
  var weekCount = 1;

  for (var d = 1; d <= daysInMonth; d++) {
    var dateObj = new Date(CONFIG.year, monthIndex, d);
    var dayName = Utilities.formatDate(dateObj, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "EEEE");
    dayName = dayName.charAt(0).toUpperCase() + dayName.slice(1);

    if (d === 1 || dayName === "Monday" || dayName === "Lunedì") { 
        var label = (d===1 && dayName!=="Monday" && dayName!=="Lunedì") ? "WEEK 1 (Partial)" : "WEEK " + weekCount;
        if((dayName === "Monday" || dayName === "Lunedì") && d !== 1) weekCount++;
        sheet.getRange(currentRow, 2, 1, 7).merge().setValue(label).setBackground(CONFIG.colors.week).setFontColor("white").setFontWeight("bold");
        currentRow++;
    }
    var range = sheet.getRange(currentRow, 2, CONFIG.rowsPerDay, 7);
    sheet.getRange(currentRow, 2, CONFIG.rowsPerDay, 1).setValue(dateObj).setNumberFormat("dd/MM");
    sheet.getRange(currentRow, 3, CONFIG.rowsPerDay, 1).setValue(dayName);
    sheet.getRange(currentRow, 6, CONFIG.rowsPerDay, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.categories).build());
    sheet.getRange(currentRow, 7, CONFIG.rowsPerDay, 1).setNumberFormat("€#,##0.00");
    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID); 
    range.setBorder(null, null, true, null, null, null, "#B0BEC5", SpreadsheetApp.BorderStyle.DOTTED); 
    if (CONFIG.rowsPerDay > 1) { sheet.getRange(currentRow + 1, 2, CONFIG.rowsPerDay - 1, 2).setFontColor("white"); }
    currentRow += CONFIG.rowsPerDay;
  }
  addWeeklyAnalysisMatrix(sheet);
  sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 110); sheet.setColumnWidth(4, 50); sheet.setColumnWidth(5, 160); sheet.setColumnWidth(6, 160); 
}

function setupTopSection(sheet) {
  sheet.getRange("B4:C4").merge().setValue("1. INCOME").setBackground(CONFIG.colors.subHeader).setFontWeight("bold");
  var incRange = sheet.getRange(LAYOUT.incomeStart, 2, CONFIG.incomeItems.length, 1);
  incRange.setValues(CONFIG.incomeItems.map(x => [x]));
  sheet.getRange(LAYOUT.incomeStart, 3, CONFIG.incomeItems.length, 1).setNumberFormat("€#,##0.00");
  sheet.getRange(LAYOUT.incomeTotalRow, 3).setFormula(`=SUM(C${LAYOUT.incomeStart}:C${LAYOUT.incomeEnd})`)
       .setBackground("#F5F5F5").setFontWeight("bold").setNumberFormat("€#,##0.00");

  sheet.getRange("E4:F4").merge().setValue("2. FIXED").setBackground(CONFIG.colors.subHeader).setFontWeight("bold");
  var fixRange = sheet.getRange(LAYOUT.fixedStart, 5, CONFIG.fixedExpenses.length, 1);
  fixRange.setValues(CONFIG.fixedExpenses.map(x => [x]));
  sheet.getRange(LAYOUT.fixedStart, 6, CONFIG.fixedExpenses.length, 1).setBackground("white").setBorder(true, true, true, true, true, true, "#B0BEC5", null).setNumberFormat("€#,##0.00");
  sheet.getRange(LAYOUT.fixedTotalRow, 6).setFormula(`=SUM(F${LAYOUT.fixedStart}:F${LAYOUT.fixedEnd})`).setBackground("#F5F5F5").setFontWeight("bold").setNumberFormat("€#,##0.00");

  sheet.getRange("H4:I4").merge().setValue("3. SAVINGS").setBackground(CONFIG.colors.subHeader).setFontWeight("bold");
  sheet.getRange("H5").setValue("Goal (Input)"); sheet.getRange("I5").setNumberFormat("€#,##0.00");
  sheet.getRange("H6").setValue("Ref (20%)").setFontStyle("italic").setHorizontalAlignment("right");
  sheet.getRange("I6").setFormula(`=C${LAYOUT.incomeTotalRow}*20%`).setNumberFormat("€#,##0.00").setFontColor("grey"); 
  
  // REMOVED: sheet.getRange("F15").setFormula("=I5")... (No longer needed)

  var mathFormula = `=C${LAYOUT.incomeTotalRow}-F${LAYOUT.fixedTotalRow}-I5`;
  var customFormat = '"AVAILABLE: "€#,##0.00'; 
  sheet.getRange("K4:L8").merge().setBackground(CONFIG.colors.header).setFontColor("white").setHorizontalAlignment("center").setVerticalAlignment("middle")
       .setFormula(mathFormula).setNumberFormat(customFormat).setFontSize(14).setFontWeight("bold").setWrap(true);

  var sRow = LAYOUT.statusRow;
  var logDataStart = LAYOUT.logStartRow + 1;
  sheet.getRange(sRow, 5).setValue("Spese Var.:").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(sRow, 6).setFormula(`=SUM(G${logDataStart}:G)`).setNumberFormat("€#,##0.00").setBackground("#F5F5F5").setFontColor("black").setFontWeight("bold");
  sheet.getRange(sRow, 7).setValue("Rimanente:").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(sRow, 8).setFormula(`=(C${LAYOUT.incomeTotalRow}-F${LAYOUT.fixedTotalRow}-I5)-F${sRow}`).setNumberFormat("€#,##0.00").setBackground(CONFIG.colors.warning).setFontColor("black").setFontWeight("bold");
}

function addWeeklyAnalysisMatrix(sheet) {
  var startRow = LAYOUT.logStartRow - 1;
  var colStart = 14; 
  var logDataStart = LAYOUT.logStartRow + 1;
  sheet.getRange(startRow, colStart, 1, 6).merge().setValue("📅 WEEKLY HEATMAP").setBackground(CONFIG.colors.header).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(startRow+1, colStart, 1, 6).setValues([["Category", "W1", "W2", "W3", "W4", "W5"]]).setBackground(CONFIG.colors.subHeader).setFontWeight("bold").setHorizontalAlignment("center");
  CONFIG.categories.forEach(function(cat, index) {
    var currentRow = startRow + 2 + index;
    sheet.getRange(currentRow, colStart).setValue(cat);
    for(var w=1; w<=5; w++) {
       var min = (w-1)*7 + 1; var max = w*7;
       var dateRange = `$B$${logDataStart}:$B`; var catRange = `$F$${logDataStart}:$F`; var amtRange = `$G$${logDataStart}:$G`;
       var formula = `=SUMPRODUCT((${catRange}="${cat}") * (IFERROR(DAY(${dateRange});0)>=${min}) * (IFERROR(DAY(${dateRange});0)<=${max}) * ${amtRange})`;
       if(w===5) formula = `=SUMPRODUCT((${catRange}="${cat}") * (IFERROR(DAY(${dateRange});0)>28) * ${amtRange})`;
       sheet.getRange(currentRow, colStart+w).setFormula(formula).setNumberFormat("€0"); 
    }
  });
  var totalRow = startRow + 2 + CONFIG.categories.length;
  sheet.getRange(totalRow, colStart).setValue("TOTAL").setFontWeight("bold");
  for(var w=1; w<=5; w++) { 
    var colLetter = String.fromCharCode(78+w); 
    var sumStart = startRow + 2; var sumEnd = totalRow - 1;
    sheet.getRange(totalRow, colStart+w).setFormula(`=SUM(${colLetter}${sumStart}:${colLetter}${sumEnd})`).setNumberFormat("€0"); 
  }
}

function createDashboard(ss) {
  var sheet = ss.getSheetByName("DASHBOARD");
  if (!sheet) sheet = ss.insertSheet("DASHBOARD");
  sheet.clear(); sheet.setHiddenGridlines(true);
  
  formatHeader(sheet.getRange("B2:H2"), `YEARLY OVERVIEW`, CONFIG.colors.header);
  var headers = [["Month", "Income", "Goal", "Fixed", "Spent", "Left", "Status"]];
  sheet.getRange("B4:H4").setValues(headers).setBackground(CONFIG.colors.subHeader).setFontWeight("bold");
  
  var row = 5;
  CONFIG.months.forEach(m => {
    sheet.getRange(row, 2).setValue(m);
    sheet.getRange(row, 3).setFormula(`='${m}'!C${LAYOUT.incomeTotalRow}`);
    
    // FIX IS HERE: Point to I5 (Input Goal), not F15
    sheet.getRange(row, 4).setFormula(`='${m}'!I5`);
    
    sheet.getRange(row, 5).setFormula(`='${m}'!F${LAYOUT.fixedTotalRow}`);
    sheet.getRange(row, 6).setFormula(`='${m}'!F${LAYOUT.statusRow}`); 
    sheet.getRange(row, 7).setFormula(`='${m}'!H${LAYOUT.statusRow}`);
    
    var total = `(F${row}+G${row})`; 
    var pct = `IF(${total}=0; 0; F${row}/${total})`; 
    var filled = `MIN(10; ROUND(${pct}*10))`;
    var empty = `MAX(0; 10-${filled})`;
    var barFormula = `=REPT("■"; ${filled}) & REPT("□"; ${empty})`; 
    sheet.getRange(row, 8).setFormula(barFormula).setHorizontalAlignment("center").setFontColor("#546E7A");
    row++;
  });
  
  sheet.getRange("C5:G17").setNumberFormat("€#,##0"); 
  createCharts(sheet);
  
  var statusRange = sheet.getRange("H5:H16");
  var ruleGreen = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G5>=0').setFontColor("#66BB6A").setRanges([statusRange]).build();
  var ruleRed = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G5<0').setFontColor("#D32F2F").setRanges([statusRange]).build();
  sheet.setConditionalFormatRules([ruleGreen, ruleRed]);
}

function createCharts(sheet) {
  // Detailed Category Table
  sheet.getRange("J5").setValue("Category").setFontWeight("bold"); 
  sheet.getRange("K5").setValue("Total").setFontWeight("bold");
  var chartDataStartRow = 6;
  var rangeEnd = chartDataStartRow + CONFIG.categories.length;

  CONFIG.categories.forEach(function(cat, i) {
    sheet.getRange(chartDataStartRow+i, 10).setValue(cat);
    var ranges = [];
    CONFIG.months.forEach(function(m) { 
       var rowInMonthly = (LAYOUT.logStartRow - 1) + 2 + i;
       ranges.push(`SUM('${m}'!O${rowInMonthly}:S${rowInMonthly})`); 
    });
    sheet.getRange(chartDataStartRow+i, 11).setFormula(`=${ranges.join("+")}`).setNumberFormat("€#,##0");
  });

  // Macro Category Table
  var macroCol = 13; 
  sheet.getRange(5, macroCol).setValue("Macro Category").setFontWeight("bold");
  sheet.getRange(5, macroCol+1).setValue("Total").setFontWeight("bold");
  
  var macros = ["Needs", "Optional", "Culture", "Extra"];
  var macroRows = 6;
  
  macros.forEach(function(macro, i) {
    sheet.getRange(macroRows+i, macroCol).setValue(macro);
    var catRange = `J$${chartDataStartRow}:J$${rangeEnd}`;
    var sumRange = `K$${chartDataStartRow}:K$${rangeEnd}`;
    var formula = `=SUMIF(${catRange}; "${macro}*"; ${sumRange})`;
    sheet.getRange(macroRows+i, macroCol+1).setFormula(formula).setNumberFormat("€#,##0");
  });
  
  var charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  var pie1 = sheet.newChart().asPieChart().addRange(sheet.getRange(5, 10, CONFIG.categories.length + 1, 2)).setPosition(19, 2, 0, 0)
    .setTitle("Detailed Spending").setOption('pieSliceText', 'label').build();
  sheet.insertChart(pie1);

  var pie2 = sheet.newChart().asPieChart().addRange(sheet.getRange(5, macroCol, 5, 2)).setPosition(19, 8, 0, 0)
    .setTitle("Macro Categories (Needs/Wants)").setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#90CAF9', '#F48FB1', '#A5D6A7', '#FFCC80']).build();
  sheet.insertChart(pie2);
}

function protectSystem(ss) {
  var sheets = ss.getSheets();
  sheets.forEach(s => {
    if (s.getName() === "DASHBOARD") return;
    var p = s.protect().setWarningOnly(true);
    var incRange = s.getRange(LAYOUT.incomeStart, 3, CONFIG.incomeItems.length, 1);
    var fixRange = s.getRange(LAYOUT.fixedStart, 6, CONFIG.fixedExpenses.length, 1);
    var savRange = s.getRange("I5");
    var logRange = s.getRange(LAYOUT.logStartRow + 1, 4, 300, 5);
    p.setUnprotectedRanges([incRange, fixRange, savRange, logRange]);
  });
}

function formatHeader(range, text, bg) {
  range.merge().setValue(text).setBackground(bg).setFontColor("white")
       .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
}