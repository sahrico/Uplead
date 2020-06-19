function loginUplead(){
  Aliases.browser.pageUpleadLogIn.formLogin.textLogin.SetText(Project.Variables.varLogin);
  Delay(500);
  
  Aliases.browser.pageUpleadLogIn.formLogin.textPassword.SetText(Project.Variables.varPassword);
  Delay(500);
  
  Aliases.browser.pageUpleadLogIn.formLogin.buttonLogin.Click();
  Delay(2000);
  
  Library.waitForPages(Aliases.browser.pageUplead);
}

function navigateUpleadPage(){
  var URL = Project.Variables.varURL;
 
  Browsers.Item(btChrome).Navigate(URL);
  Delay(500);
  
}

function navigateCompanySearchPage(){
  var URLCompanySearch = Project.Variables.varURLCompanySearch;
  
  Browsers.Item(btChrome).Navigate(URLCompanySearch);
  
  Delay(1000);
  
}
function closeChromeBrowser(){
  
  var p = Sys.Find("ProcessName", "chrome");
  
  while(p.Exists){
    p.Close();
  }  
}

function logoutUplead(){
  Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panelLogout.panel.panel.panel.setaLogout.Click();
  Delay(500);
  
  Aliases.browser.pageUpleadCompanySearch.panelRoot.linkLogout.Click();
  Delay(2000);
  
  Library.waitForPages(Aliases.browser.pageUpleadLogIn);
}

function searchCompany(COMPANY){
  
  navigateCompanySearchPage();
//  Aliases.browser.pageUplead.linkCompanySearch.Click();
//  Delay(2000);
  
  Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.header.textQuickSearch.SetText(COMPANY);
  Delay(1200);
  
  Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.header.textQuickSearch.Keys("[Enter]");
  Delay(2500);
}

function waitForPages(pagePath){
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function is used when a page takes more time than usual to load. Waits until the page is completed loaded.
	 * 
	 * @Usage Example: waitForPages(Aliases.browser.pageHomeSalesforce);
	 * 
	 * @Special Note:
*/

  for(var i = 0 ; i < 20 ; i++){
    if(pagePath.Exists){
      i=20;
    }else{
      Delay(1000);
    }
  }
}

function getCurrentDate(){

/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function gets and returns the current date.
	 * 
	 * @Usage Example: getCurrentDate();
	 * 
	 * @Special Note:
*/
  var today = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%m/%d/%Y"); 
  return today;
}

function getCurrentTime(){
  
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function gets and returns the current time (11:02:00).
	 * 
	 * @Usage Example: getCurrentTime();
	 * 
	 * @Special Note:
*/

  var time = aqConvert.DateTimeToFormatStr(aqDateTime.Time(), "%H:%M:%S");
  return time;
}

function getLogFileName(){  
  
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function sets and returns the full log file name, that is a union between date, time and name (04282020_110710_LogFile.xlsx).
	 * 
	 * @Usage Example: getLogFileName();
	 * 
	 * @Special Note:
*/

  //Replacing Special Caracteres
  date = ReplaceSpChr(getCurrentDate());
  time = ReplaceSpChr(getCurrentTime());
  
  //Setting the file name
  var sheetName = date+"_"+time+"_"+"UpdatedCompanies.xlsx";
      
  return sheetName;
}

function ReplaceSpChr(StringVal){

/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function is used to remove special caracteres of a string.
	 * 
	 * @Usage Example: ReplaceSpChr("10/07/2020");
   *                 Returns "10072020"
	 * 
	 * @Special Note:
*/
    // Create regular expression pattern.
    var re = /\D/g;
    
    // Use replace
    var r = StringVal.replace(re, "");
    
    return(r);
}

function closeExcelProcess(){
  
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function closes the Excel Process.
	 * 
	 * @Usage Example: closeExcelProcess();
	 * 
	 * @Special Note:
*/

  
  var p = Sys.Find("ProcessName", "EXCEL");
  
  while(p.Exists){
     p.Terminate();
  }
}


function createExcelLogFile(){
  
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function is used to create a excel log file.
	 * 
	 * @Usage Example: createExcelLogFile();
	 * 
	 * @Special Note:
*/

  //Close excel process
  closeExcelProcess();
  
  //Get the excel application
  var excelApp = Sys.OleObject("Excel.Application");
  
  //Add a new spreadsheet
  var spreadSheet = excelApp.Workbooks.Add();
  
  //Get the spreadsheet
  var currentSheet = spreadSheet.ActiveSheet;
  currentSheet.Visible = true;

  //Setting the head columns
  currentSheet.Cells.Item(1, 1).Value2 = "Company Email";
  currentSheet.Cells.Item(1, 2).Value2 = "Company Name";
  currentSheet.Cells.Item(1, 3).Value2 = "Company Industry";
  
  //Setting the spreadsheet name
  sheetName = getLogFileName();
  
  //Setting the path where the spreadsheet will be saved
  path = Project.Path+Project.Variables.logFilePath;
  
  //Setting the final path
  var filePath = path+sheetName;
  
  //Setting the new value of the file path
  Project.Variables.logFilePath = filePath;
 
  //Save excel file and close   
  currentSheet.SaveAs(filePath);
  excelApp.Quit();
}

function writeExcelLogFile(companyEmail,companyName,companyIndustry)
{
  
/*
	 * @Author: Samantha Rico
	 * 
	 * @Creation date: 27-Mar-2020
	 * 
	 * @Update author:
	 * 
	 * @Last update: 
	 * 
	 * @Abstract: This function is used to write some data into a excel log file.
	 * 
	 * @Usage Example: writeExcelLogFile("email@email.com","DELETED","It was deleted because it is a duplicated lead");
	 * 
	 * @Special Note:
*/

  //Close Excel Process
  closeExcelProcess();
  
  //Set File Name  
  var FileName = Project.Variables.logFilePath;
  
  //Get Excel Object
  let Excel = getActiveXObject("Excel.Application");

  //Open the Excel File
  Excel.Workbooks.Open(FileName);
  
  //Disable Alerts
  Excel.Application.DisplayAlerts = false;
  Excel.Application.AlertBeforeOverwriting = false; 
  
  //Row and Columns Count  
  let RowCount = Excel.ActiveSheet.UsedRange.Rows.Count;
  let ColumnCount = Excel.ActiveSheet.UsedRange.Columns.Count;
    
  //Inserting the data into the spreadsheet
  for(let i = 1 ; i<=ColumnCount; i++){
    switch (i) {
      case 1:
         Excel.Cells.Item(RowCount+1, i).Value2 = companyEmail;
      break;
      case 2:
         Excel.Cells.Item(RowCount+1, i).Value2 = companyName;
      break;
      case 3:
         Excel.Cells.Item(RowCount+1, i).Value2 = companyIndustry;
      break;
    }
  }
    
  //Saving the file
  Excel.Application.ActiveWorkbook.SaveAs(FileName);

  //Quit the excel file
  Excel.Quit();
  
  //Close Excel Process
  closeExcelProcess();
}


module.exports.loginUplead = loginUplead;
module.exports.closeChromeBrowser = closeChromeBrowser;
module.exports.searchCompany = searchCompany;
module.exports.logoutUplead = logoutUplead;
module.exports.createExcelLogFile = createExcelLogFile;
module.exports.writeExcelLogFile = writeExcelLogFile;
module.exports.navigateUpleadPage = navigateUpleadPage;
module.exports.navigateCompanySearchPage = navigateCompanySearchPage;