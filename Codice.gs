//GSHEET formula beautifier:
//generated from https://itectec.com/webapp/google-sheets-pretty-print-google-sheet-formula/
/*

███████╗ ██████╗ ██████╗ ███╗   ███╗██╗   ██╗██╗      █████╗         
██╔════╝██╔═══██╗██╔══██╗████╗ ████║██║   ██║██║     ██╔══██╗        
█████╗  ██║   ██║██████╔╝██╔████╔██║██║   ██║██║     ███████║        
██╔══╝  ██║   ██║██╔══██╗██║╚██╔╝██║██║   ██║██║     ██╔══██║        
██║     ╚██████╔╝██║  ██║██║ ╚═╝ ██║╚██████╔╝███████╗██║  ██║        
╚═╝      ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝        
                                                                     
██████╗ ███████╗ █████╗ ██╗   ██╗████████╗███████╗██╗███████╗██████╗ 
██╔══██╗██╔════╝██╔══██╗██║   ██║╚══██╔══╝██╔════╝██║██╔════╝██╔══██╗
██████╔╝█████╗  ███████║██║   ██║   ██║   █████╗  ██║█████╗  ██████╔╝
██╔══██╗██╔══╝  ██╔══██║██║   ██║   ██║   ██╔══╝  ██║██╔══╝  ██╔══██╗
██████╔╝███████╗██║  ██║╚██████╔╝   ██║   ██║     ██║███████╗██║  ██║
╚═════╝ ╚══════╝╚═╝  ╚═╝ ╚═════╝    ╚═╝   ╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝



*/
function onInstall() {
  onOpen();
}

function onOpen() {
  ShowMyMenu();
}

function ShowMyMenu() {

  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu();
  menu
    .addItem('Formula Updater', 'OpenFormula')
    .addToUi();

}


/**
 * SHOW DIALOG BOX
 * 
 */
function OpenFormula() {
  //var cell = SpreadsheetApp.getCurrentCell();
  //var formula = cell.getFormula();
  var ui = SpreadsheetApp.getUi();

  var template = HtmlService.createTemplateFromFile('OpenFormula');

  // get output html
  var html = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(1200)
    .setHeight(1000);

  // show modal window
  ui.showModalDialog(html, 'Update formula:');
}

/**
 * Retrieve Formula from the curent Cell
 * 
 */
function getFormula() {
  var cell = SpreadsheetApp.getCurrentCell();
  var formula = cell.getFormula();
  return formula;
}


function setFormula(formula) {
  var cell = SpreadsheetApp.getCurrentCell();
  cell.offset(0, 0).setFormula(formula);
}



/**
 * saveMyFormula
 */

function saveMyFormula(formula, comments) {

  var temp = new Date();
  var mydate = Utilities.formatDate(temp, "CET", "yyyy-MM-dd HH-mm"); //"yyyy-MM-dd"); v1.1
  var cell = SpreadsheetApp.getCurrentCell();
  var cellNotation = cell.getA1Notation();
  var filename = SpreadsheetApp.getActive().getName();
  var actualSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var NewFileName = mydate + " : " + "formula:" + actualSheetName + "!" + cellNotation + ' - File: ' + filename;
  savemyformula(formula, NewFileName, createOrGetFolder("Backup_formulas", getParentFolder()));
  //setting formula & note
  var currentnote = cell.offset(0, 0).getNote();
  cell.offset(0, 0).setFormula(formula);
  cell.offset(0, 0).setNote("--> " + actualSheetName + "!" + cellNotation +" "+"Updated at "+mydate+". " + "      "+comments+ "    --------------------------  "+currentnote);

}


function beautify() {
  var cell = SpreadsheetApp.getCurrentCell();
  var formula = cell.getFormula();
  //formula= minify(formula);

  //formula = formula.replace(/\s+/g, ""); //substitute all spaces
  formula = formula.replace(/[;]\s+/gm, ";");
  formula = formula.replace(/[(]\s+/gm, "(");
  formula = formula.replace(/[)]\s+/gm, ")");
  //cell.offset(0, 0).setFormula(prettify(formula));
  return prettify(formula);
}

function minify(formula) {
  var cell = SpreadsheetApp.getCurrentCell();
  var formula = cell.getFormula();
  formula = formula.replace(/\s+/g, ""); //substitute all spaces
  /*
  // 
  formula = formula.replace(/[;]\s+/gm, ";");
  formula = formula.replace(/[(]\s+/gm, "(");
  formula = formula.replace(/[)]\s+/gm, ")");
  formula = formula.replace(/\s+[)]/gm, ")");
  */
  //cell.offset(0, 0).setFormula(formula);
  return minify;
}

function prettify() {
  var cell = SpreadsheetApp.getCurrentCell();
  var formula = cell.getFormula();
  var pretty = '';
  var tabNum = 1;
  var tabOffset = 4;
  var tabs = [];
  const regexif = /[\{\(]/;     //    /[\{\(]/   /SE[(]/gm    //https://stackoverflow.com/questions/56762554/what-is-the-regexmatch-equivalent-in-google-apps-script

  formula.split('').forEach(function (c, i) {
    if (regexif.test(c)) {    //       //    /[\{\(]/
      tabNum++;
      tabs[tabNum] = (tabs[tabNum - 1] ? tabs[tabNum - 1] : 0) + tabOffset + 1;
      //tabOffset = 0;
      pretty += c + '\n' + ' '.repeat(tabs[tabNum]);
    } else if (/[\}\)]/.test(c)) {
      tabNum--;
      pretty += c + '\n' + ' '.repeat(tabs[tabNum]);
      //tabOffset = 0;
    } else if (/[,;]/.test(c)) {

      pretty += c + '\n' + ' '.repeat(tabs[tabNum]);
      //tabOffset = 0;
    } else {
      pretty += c;
      //tabOffset++;
    }
  });
  return pretty;
}



