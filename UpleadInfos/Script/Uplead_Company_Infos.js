﻿var Library = require("Library");


function upleadCompanyInfos(databankPath){
  
  //Open the driver
  var databankCompany = DDT.CSVDriver(databankPath);
  
  Library.closeChromeBrowser();
  
  //Create Log File Excel
  Library.createExcelLogFile();
  
  Library.navigateUpleadPage();
  
  //Checks if the logout option exists
  if(Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panelLogout.panel.panel.panel.setaLogout.Exists)
  {
    Library.logoutUplead();
    Library.loginUplead();
  }else{
    Library.loginUplead();
  } 
  
  count = 0;
  
  //Data Driven Loop
  while(!databankCompany.EOF()){
    
    count++;
    
    companyName = "";
    companyIndustry = "";
    companyEmail = "";
  
    companyEmail = databankCompany.Value(0);
    
    Log.AppendFolder(count+". "+companyEmail);
    
    Library.searchCompany(companyEmail);
    
    Log.Picture(Aliases.browser.pageUpleadCompanySearch,"Screenshot of the Search Result!");
        
    if(Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.linkCompanyFounded.Exists){
    
      //It clicks on the company returned
      Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.linkCompanyFounded.Click();
      Delay(2000);
    
      //Get Company Name
      companyName = Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.panelCompanyName.panel.panel.panel.panel.panel.panel.panel.headerCompanyFounded.companyName.contentText;
      companyIndustry = Aliases.browser.pageUpleadCompanySearch.panelRoot.panel.panel.panelIndustry.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.panel.companyIndustry.contentText;

      Log.Message("Company Name: "+companyName);
      Log.Message("Company Industry: "+companyIndustry);
      
      Log.Picture(Aliases.browser.pageUpleadCompanySearch, "Screenshot of the Company Page!");
      
    }else{
      Log.Message("The company was not found");
      companyName = "Not Found!";
      
    }//End Else
    
    Library.writeExcelLogFile(companyEmail,companyName,companyIndustry);
   
    Log.PopLogFolder();
    
    databankCompany.Next();
  }//End While
  
    //Close driver
    DDT.CloseDriver(databankCompany.Name);
    
    Library.logoutUplead();
    Library.closeChromeBrowser(); 
}