// Granular Anomaly Detector Script
//
// Copyright 2016 - Optmyzr Inc - All Rights Reserved
// Visit www.optmyzr.com for more AdWords Scripts and PPC Management Tools and Reports
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

currentSetting = {};
var DEBUG = 0;
var VERBOSE = 1;

function main(){
  
  
  // UPDATE THESE SETTINGS
  var accountName = "CHCP"; //currentSetting['accountName'];
  var currentPeriodStartsNDaysAgo = 7;//1
  var currentPeriodEndsNDaysAgo = 1; 
  var previousPeriodStartsNDaysAgo = 14;//8
  var previousPeriodEndsNDaysAgo = 8; 
  
  var includeAccountLevel = 0;//= currentSetting['includeAccountLevel'];
  var includeCampaignLevel = 1;//= currentSetting['includeCampaignLevel'];
  var includeAdGroupLevel = 1;//= currentSetting['includeAdGroupLevel'];
  var includeKeywordLevel = 0;//= currentSetting['includeKeywordLevel'];
  var includeAdLevel = 0;//= currentSetting['includeAdLevel'];
  
  currentSetting.minAlertAllConversions = 20;
  currentSetting.minDecreaseForAllConversionsAlert = 0;
  currentSetting.minIncreaseForAllConversionsAlert = 0.1;
  currentSetting.minAlertAllConversionValue = 20;
  currentSetting.minDecreaseForAllConversionValueAlert = 0;
  currentSetting.minIncreaseForAllConversionValueAlert = 0.1;
  currentSetting.minAlertAllConversionRate = 20;
  currentSetting.minDecreaseForAllConversionRateAlert = 0;
  currentSetting.minIncreaseForAllConversionRateAlert = 0.1;
  currentSetting.minAlertAverageCpc = 20;
  currentSetting.minDecreaseForAverageCpcAlert = 0;
  currentSetting.minIncreaseForAverageCpcAlert = 0.1;
  currentSetting.minAlertCtr = 20;
  currentSetting.minDecreaseForCtrAlert = 0;
  currentSetting.minIncreaseForCtrAlert = 0.1;
  currentSetting.minAlertImpressions = 20;
  currentSetting.minDecreaseForImpressionsAlert = 0;
  currentSetting.minIncreaseForImpressionsAlert = 0.1;
  currentSetting.minAlertClicks = 20;
  currentSetting.minDecreaseForClicksAlert = 0;
  currentSetting.minIncreaseForClicksAlert = 0.1;
  currentSetting.minAlertAveragePosition = 20;
  currentSetting.minDecreaseForAveragePositionAlert = 0;
  currentSetting.minIncreaseForAveragePositionAlert = 0.1;
  currentSetting.minAlertCost = 20;
  currentSetting.minDecreaseForCostAlert = 0;
  currentSetting.minIncreaseForCostAlert = 0.1;
  currentSetting.minAlertConversionRate = 20;
  currentSetting.minDecreaseForConversionRateAlert = 0;
  currentSetting.minIncreaseForConversionRateAlert = 0.1;
  currentSetting.minAlertConversions = 20;
  currentSetting.minDecreaseForConversionsAlert = 0;
  currentSetting.minIncreaseForConversionsAlert = 0.1;
  currentSetting.minAlertConversionValue = 20;
  currentSetting.minDecreaseForConversionValueAlert = 0;
  currentSetting.minIncreaseForConversionValueAlert = 0.1;
  //currentSetting.minAlertConvertedClicks = 20;
  //currentSetting.minDecreaseForConvertedClicksAlert = 0;
  //currentSetting.minIncreaseForConvertedClicksAlert = 0.1;
  currentSetting.minAlertCostPerConversion = 20;
  currentSetting.minDecreaseForCostPerConversionAlert = 0;
  currentSetting.minIncreaseForCostPerConversionAlert = 0.1;
  currentSetting.minAlertCostPerAllConversion = 20;
  currentSetting.minDecreaseForCostPerAllConversionAlert = 0;
  currentSetting.minIncreaseForCostPerAllConversionAlert = 0.1;
  currentSetting.minAlertCrossDeviceConversions = 20;
  currentSetting.minDecreaseForCrossDeviceConversionsAlert = 0;
  currentSetting.minIncreaseForCrossDeviceConversionsAlert = 0.1;
  currentSetting.minAlertValuePerConversion = 20;
  currentSetting.minDecreaseForValuePerConversionAlert = 0;
  currentSetting.minIncreaseForValuePerConversionAlert = 0.1;
  currentSetting.minAlertValuePerAllConversion = 20;
  currentSetting.minDecreaseForValuePerAllConversionAlert = 0;
  currentSetting.minIncreaseForValuePerAllConversionAlert = 0.1;
  currentSetting.minAlertViewThroughConversions = 20;
  currentSetting.minDecreaseForViewThroughConversionsAlert = 0;
  currentSetting.minIncreaseForViewThroughConversionsAlert = 0.1;

  /*change where the email goes here */
  
  currentSetting.accountManagers = "digitalad@agency451.com, digitalad@agency451.com";
  currentSetting.emailAddresses = "digitalad@agency451.com, digitalad@agency451.com";
  // END OF SETTINGS
  

      
  includeAccountLevel = (includeAccountLevel==true|| includeAccountLevel=="true")?true:false;
  includeCampaignLevel = (includeCampaignLevel==true|| includeCampaignLevel=="true")?true:false;
  includeAdGroupLevel = (includeAdGroupLevel==true|| includeAdGroupLevel=="true")?true:false;
  includeKeywordLevel = (includeKeywordLevel==true|| includeKeywordLevel=="true")?true:false;
  includeAdLevel = (includeAdLevel==true|| includeAdLevel=="true")?true:false;
   
  
  // Script Variables
  var spreadsheetUrl = "new"; //currentSetting['spreadsheetUrl'];
  var time = "Days Ago"; //currentSetting['Time'];
  var campaignNameSelectorStatement = "CampaignName CONTAINS_IGNORE_CASE ''"; //= currentSetting['CampainNameSelectorStatement'];
  var segments = ""; //= currentSetting['segments'];  // AdNetworkType1, AdNetworkType2, ClickType, DayOfWeek, Device, MonthOfYear, Slot, 
  
  var decimalPoint = ".";
  var debug = 0;
  var numFrozenRows = 2;
  var alphabet = new Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ");
  var elements = new Array();
  
  var accountAlertCount = 0;
  var campaignAlertCount = 0;
  var adGroupAlertCount = 0;
  var keywordAlertCount = 0;
  var adAlertCount = 0;
  
  var accountAttributeColumns = ['AccountDescriptiveName'];
  var campaignAttributeColumns = ['AccountDescriptiveName', 
                                  'CampaignName'];
  var adGroupAttributeColumns = ['AccountDescriptiveName', 
                                 'CampaignName',
                                 'AdGroupName'];
  var keywordAttributeColumns = ['AccountDescriptiveName', 
                                 'CampaignName',
                                 'AdGroupName',
                                 'Criteria'];
  var adAttributeColumns = ['AccountDescriptiveName', 
                            'CampaignName',
                            'AdGroupName',
                            'Headline',
                            'Description1',
                            'Description2',
                            'DisplayUrl',
                            'Id'];
  var metricsColumns = ['Conversions',
                        'Clicks', 
                        'Cost',
                        'Impressions',
                        

                        'AllConversionRate',
                        'AverageCpc',
                        'Ctr', 
                        'AveragePosition', 
                        'ConversionRate',
                        'AllConversions',
                        'ConversionValue', 
                      //  'ConvertedClicks',
                        'CostPerConversion',
                        'CostPerAllConversion',
                        'CrossDeviceConversions',
                        'ValuePerConversion',
                        'ValuePerAllConversion',
                        'ViewThroughConversions'];
   /*   not in use                      
  var metricsColumns = ['AllConversions',
                        'AllConversionValue',
                        'AllConversionRate',
                        'AverageCpc',
                        'Ctr', 
                        'Impressions',

                        'Clicks', 
                        'AveragePosition', 
                        'Cost', 
                        'ConversionRate',
                        'Conversions',
                        'ConversionValue', 
                      //  'ConvertedClicks',
                        'CostPerConversion',
                        'CostPerAllConversion',
                        'CrossDeviceConversions',
                        'ValuePerConversion',
                        'ValuePerAllConversion',
                        'ViewThroughConversions'];
*/

  var metricsColumnsTrim = 4;
  
  //if statement ends with contains_ignore_case, add ''.
  if(campaignNameSelectorStatement.trim().indexOf("CONTAINS_IGNORE_CASE", campaignNameSelectorStatement.length - "CONTAINS_IGNORE_CASE".length)!=-1){
    campaignNameSelectorStatement = campaignNameSelectorStatement.trim()+ " \'\' ";
  }
  
  segments = segments.replace(" ", "");
  if(segments == "") {
    var segmentsToInclude = new Array();
  } else if(segments.trim().split(",").length > 0) {
    var segmentsToInclude = segments.split(",");
  } else if(segments.trim().split(",").length == 0) {
    var segmentsToInclude = new Array(segments);
  }

  var goodTextColor = "green";
  var badTextColor = "red";
  var goodCellColor = "#d9ffcc";
  var badCellColor = "#ffcccc";
  
  currentSetting.colors = [];
  currentSetting.colors["AllConversions"] = {};
  currentSetting.colors["AllConversions"].textColorForDecreases = badTextColor;
  currentSetting.colors["AllConversions"].textColorForIncreases = goodTextColor;
  currentSetting.colors["AllConversions"].bgColorForDecreases = badCellColor;
  currentSetting.colors["AllConversions"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["AllConversionValue"] = {};
  currentSetting.colors["AllConversionValue"].textColorForDecreases = badTextColor;
  currentSetting.colors["AllConversionValue"].textColorForIncreases = goodTextColor;
  currentSetting.colors["AllConversionValue"].bgColorForDecreases = badCellColor;
  currentSetting.colors["AllConversionValue"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["AllConversionRate"] = {};
  currentSetting.colors["AllConversionRate"].textColorForDecreases = badTextColor;
  currentSetting.colors["AllConversionRate"].textColorForIncreases = goodTextColor;
  currentSetting.colors["AllConversionRate"].bgColorForDecreases = badCellColor;
  currentSetting.colors["AllConversionRate"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["Ctr"] = {};
  currentSetting.colors["Ctr"].textColorForDecreases = badTextColor;
  currentSetting.colors["Ctr"].textColorForIncreases = goodTextColor;
  currentSetting.colors["Ctr"].bgColorForDecreases = badCellColor;
  currentSetting.colors["Ctr"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["Impressions"] = {};
  currentSetting.colors["Impressions"].textColorForDecreases = badTextColor;
  currentSetting.colors["Impressions"].textColorForIncreases = goodTextColor;
  currentSetting.colors["Impressions"].bgColorForDecreases = badCellColor;
  currentSetting.colors["Impressions"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["Clicks"] = {};
  currentSetting.colors["Clicks"].textColorForDecreases = badTextColor;
  currentSetting.colors["Clicks"].textColorForIncreases = goodTextColor;
  currentSetting.colors["Clicks"].bgColorForDecreases = badCellColor;
  currentSetting.colors["Clicks"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["ConversionRate"] = {};
  currentSetting.colors["ConversionRate"].textColorForDecreases = badTextColor;
  currentSetting.colors["ConversionRate"].textColorForIncreases = goodTextColor;
  currentSetting.colors["ConversionRate"].bgColorForDecreases = badCellColor;
  currentSetting.colors["ConversionRate"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["Conversions"] = {};
  currentSetting.colors["Conversions"].textColorForDecreases = badTextColor;
  currentSetting.colors["Conversions"].textColorForIncreases = goodTextColor;
  currentSetting.colors["Conversions"].bgColorForDecreases = badCellColor;
  currentSetting.colors["Conversions"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["ConversionValue"] = {};
  currentSetting.colors["ConversionValue"].textColorForDecreases = badTextColor;
  currentSetting.colors["ConversionValue"].textColorForIncreases = goodTextColor;
  currentSetting.colors["ConversionValue"].bgColorForDecreases = badCellColor;
  currentSetting.colors["ConversionValue"].bgColorForIncreases = goodCellColor;
  



  
  currentSetting.colors["CrossDeviceConversions"] = {};
  currentSetting.colors["CrossDeviceConversions"].textColorForDecreases = badTextColor;
  currentSetting.colors["CrossDeviceConversions"].textColorForIncreases = goodTextColor;
  currentSetting.colors["CrossDeviceConversions"].bgColorForDecreases = badCellColor;
  currentSetting.colors["CrossDeviceConversions"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["ValuePerConversion"] = {};
  currentSetting.colors["ValuePerConversion"].textColorForDecreases = badTextColor;
  currentSetting.colors["ValuePerConversion"].textColorForIncreases = goodTextColor;
  currentSetting.colors["ValuePerConversion"].bgColorForDecreases = badCellColor;
  currentSetting.colors["ValuePerConversion"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["ValuePerAllConversion"] = {};
  currentSetting.colors["ValuePerAllConversion"].textColorForDecreases = badTextColor;
  currentSetting.colors["ValuePerAllConversion"].textColorForIncreases = goodTextColor;
  currentSetting.colors["ValuePerAllConversion"].bgColorForDecreases = badCellColor;
  currentSetting.colors["ValuePerAllConversion"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["ViewThroughConversions"] = {};
  currentSetting.colors["ViewThroughConversions"].textColorForDecreases = badTextColor;
  currentSetting.colors["ViewThroughConversions"].textColorForIncreases = goodTextColor;
  currentSetting.colors["ViewThroughConversions"].bgColorForDecreases = badCellColor;
  currentSetting.colors["ViewThroughConversions"].bgColorForIncreases = goodCellColor;
  
  currentSetting.colors["AverageCpc"] = {};
  currentSetting.colors["AverageCpc"].textColorForDecreases = goodTextColor;
  currentSetting.colors["AverageCpc"].textColorForIncreases = badTextColor;
  currentSetting.colors["AverageCpc"].bgColorForDecreases = goodCellColor;
  currentSetting.colors["AverageCpc"].bgColorForIncreases = badCellColor;
  
  currentSetting.colors["AveragePosition"] = {};
  currentSetting.colors["AveragePosition"].textColorForDecreases = goodTextColor;
  currentSetting.colors["AveragePosition"].textColorForIncreases = badTextColor;
  currentSetting.colors["AveragePosition"].bgColorForDecreases = goodCellColor;
  currentSetting.colors["AveragePosition"].bgColorForIncreases = badCellColor;
  
  currentSetting.colors["Cost"] = {};
  currentSetting.colors["Cost"].textColorForDecreases = goodTextColor;
  currentSetting.colors["Cost"].textColorForIncreases = badTextColor;
  currentSetting.colors["Cost"].bgColorForDecreases = goodCellColor;
  currentSetting.colors["Cost"].bgColorForIncreases = badCellColor;
  
  currentSetting.colors["CostPerConversion"] = {};
  currentSetting.colors["CostPerConversion"].textColorForDecreases = goodTextColor;
  currentSetting.colors["CostPerConversion"].textColorForIncreases = badTextColor;
  currentSetting.colors["CostPerConversion"].bgColorForDecreases = goodCellColor;
  currentSetting.colors["CostPerConversion"].bgColorForIncreases = badCellColor;
  currentSetting.colors["CostPerAllConversion"] = {};
  currentSetting.colors["CostPerAllConversion"].textColorForDecreases = goodTextColor;
  currentSetting.colors["CostPerAllConversion"].textColorForIncreases = badTextColor;
  currentSetting.colors["CostPerAllConversion"].bgColorForDecreases = goodCellColor;
  currentSetting.colors["CostPerAllConversion"].bgColorForIncreases = badCellColor;
  
  
  var date = new Date();
  preStartDate = new Date(date.getTime()-(previousPeriodStartsNDaysAgo*24 * 60 * 60 * 1000));
  preEndDate = new Date(date.getTime()-(previousPeriodEndsNDaysAgo*24 * 60 * 60 * 1000));
  postStartDate = new Date(date.getTime()-(currentPeriodStartsNDaysAgo *24 * 60 * 60 * 1000));
  postEndDate = new Date(date.getTime()-(currentPeriodEndsNDaysAgo *24 * 60 * 60 * 1000));
  
  var preStart = preStartDate.yyyymmdd();
  var preEnd = preEndDate.yyyymmdd();
  var postStart= postStartDate.yyyymmdd();
  var postEnd = postEndDate.yyyymmdd();
  if(VERBOSE) Logger.log("Date ranges used for comparison: " + postStart + "-" + postEnd + " vs. " + preStart + "-" + preEnd); 
  
  var emailBody = "There are anomaly alerts for your AdWords account when comparing the performance for " + postStart + "-" + postEnd + " with " + preStart + "-" + preEnd;
  emailBody += "<ul>";
  
  // SET UP SPREADSHEET
  var overWriteOldData = 1;
  var targetFolder = "";
  var sheetNames = "";
  var spreadsheetName = "AdWords Alerts for account " + AdWordsApp.currentAccount().getName() + " (" + AdWordsApp.currentAccount().getCustomerId() + ") " + postStart + "-" + postEnd + " vs. " + preStart + "-" + preEnd;
  var destinationSpreadsheet = setUpReportInGoogleSheets(spreadsheetUrl, spreadsheetName, currentSetting.accountManagers, overWriteOldData, sheetNames, targetFolder)
  var spreadsheetUrl = destinationSpreadsheet.url;
  var spreadsheet = destinationSpreadsheet.spreadsheet;
  
  if(includeAccountLevel === true) {
    var acctSheet = spreadsheet.insertSheet("Account");
    
    acctSheet.getRange("AA1:CA1").clear();
    for(var i = 0; i < metricsColumns.length; i++) {
      var startColNum = i*3 + 8;
      var endColNum = i*3 + 10;
      var startCol = alphabet[startColNum] + "1";
      var endCol = alphabet[endColNum] + "1";
      var rangeText = startCol + ":" + endCol;
      var metricName = metricsColumns[i];
      acctSheet.getRange(rangeText).mergeAcross().setValue(metricName).setFontSize(10);
    }
    var startCol = alphabet[startColNum+3] + "1";
    //acctSheet.getRange(startCol).setValue("Scope").setFontSize(8);
    
    acctSheet.setFrozenRows(numFrozenRows);
    
    
    var headerValues = ["Account", "Campaign",  "Ad Group", "Keyword",
                         "Segment 1", "Segment 2", "Segment 3", "Segment 4"];
    for(var i = 0; i < metricsColumns.length; i++) {
      headerValues.push(preStart + "-" + preEnd, postStart + "-" + postEnd, "Change");
    }
   // headerValues.push("Scope");
    acctSheet.appendRow(headerValues);
    acctSheet.getRange("A1:CA1").setFontWeight("bold");
    
    
    var numColsToHide = 4 - segmentsToInclude.length;
    var segmentColIndex = 5 + segmentsToInclude.length;
    acctSheet.hideColumns(segmentColIndex, numColsToHide);
    acctSheet.hideColumns(2, 3);
    acctSheet.setFrozenColumns(8);
    acctSheet.setFrozenRows(2);
  }
  
  if(includeCampaignLevel === true) {
    var campaignSheet = spreadsheet.insertSheet("Campaigns");

    campaignSheet.getRange("AA1:CA1").clear();

    /*change to trim columns unused
    for(var i = 0; i < metricsColumns.length; i++) {
    */
    for(var i = 0; i < metricsColumnsTrim; i++) {
      var startColNum = i*3 + 8;
      var endColNum = i*3 + 10;
      var startCol = alphabet[startColNum] + "1";
      var endCol = alphabet[endColNum] + "1";
      var rangeText = startCol + ":" + endCol;
      var metricName = metricsColumns[i];
      campaignSheet.getRange(rangeText).mergeAcross().setValue(metricName).setFontSize(10);
    }
    var startCol = alphabet[startColNum+3] + "1";
   // campaignSheet.getRange(startCol).setValue("Scope").setFontSize(8);
    
    campaignSheet.setFrozenRows(numFrozenRows);
    
    var headerValues = ["Account", "Campaign", "Ad Group", "Keyword",
                         "Segment 1", "Segment 2", "Segment 3", "Segment 4"];

/*    for(var i = 0; i < metricsColumns.length; i++) {*/

    for(var i = 0; i < metricsColumnsTrim; i++) {
      headerValues.push(preStart + "-" + preEnd, postStart + "-" + postEnd, "Change");
    }
  //  headerValues.push("Scope");
    campaignSheet.appendRow(headerValues);
    campaignSheet.getRange("A1:CA1").setFontWeight("bold");
    
    var numColsToHide = 4 - segmentsToInclude.length;
    var segmentColIndex = 5 + segmentsToInclude.length;
    campaignSheet.hideColumns(segmentColIndex, numColsToHide);
    campaignSheet.hideColumns(3, 2);
    campaignSheet.setFrozenColumns(2);
  }
  
  if(includeAdGroupLevel === true) {
    var adGroupSheet = spreadsheet.insertSheet("Ad Groups");
  
  adGroupSheet.getRange("AA1:CA1").clear();

/* change made to remove headers when we aren't using columns
 for(var i = 0; i < metricsColumns.length; i++) {
 */
    for(var i = 0; i < metricsColumnsTrim; i++) {
      var startColNum = i*3 + 8;
      var endColNum = i*3 + 10;
      var startCol = alphabet[startColNum] + "1";
      var endCol = alphabet[endColNum] + "1";
      var rangeText = startCol + ":" + endCol;
      var metricName = metricsColumns[i];

      adGroupSheet.getRange(rangeText).mergeAcross().setValue(metricName).setFontSize(10);

    }
    var startCol = alphabet[startColNum+3] + "1";
   // adGroupSheet.getRange(startCol).setValue("Scope").setFontSize(8);
    
    
    
    
    
    adGroupSheet.setFrozenRows(numFrozenRows);
    var headerValues = ["Account", "Ad Group", "Campaign",  "Keyword",
                         "Segment 1", "Segment 2", "Segment 3", "Segment 4"];
    
  //  for(var i = 0; i < metricsColumns.length; i++) {
    for(var i = 0; i < metricsColumnsTrim; i++) {
      headerValues.push(preStart + "-" + preEnd, postStart + "-" + postEnd, "Change");
    }
  //  headerValues.push("Scope");
    adGroupSheet.appendRow(headerValues);
    
    adGroupSheet.getRange("A1:CA1").setFontWeight("bold");
    var numColsToHide = 4 - segmentsToInclude.length;
    var segmentColIndex = 5 + segmentsToInclude.length;
    adGroupSheet.hideColumns(segmentColIndex, numColsToHide);
    adGroupSheet.hideColumns(4, 1);
    adGroupSheet.setFrozenColumns(3);

  }
  
  if(includeKeywordLevel === true) {
    var keywordSheet = spreadsheet.insertSheet("Keywords");
    
    keywordSheet.getRange("AA1:CA1").clear();
    for(var i = 0; i < metricsColumns.length; i++) {
      var startColNum = i*3 + 8;
      var endColNum = i*3 + 10;
      var startCol = alphabet[startColNum] + "1";
      var endCol = alphabet[endColNum] + "1";
      var rangeText = startCol + ":" + endCol;
      var metricName = metricsColumns[i];
      keywordSheet.getRange(rangeText).mergeAcross().setValue(metricName).setFontSize(10);
    }
    var startCol = alphabet[startColNum+3] + "1";
   // keywordSheet.getRange(startCol).setValue("Scope").setFontSize(8);
    
    keywordSheet.setFrozenRows(numFrozenRows);
    
    var headerValues = ["Account", "Campaign", "Ad Group", "Keyword",
                         "Segment 1", "Segment 2", "Segment 3", "Segment 4"];
    for(var i = 0; i < metricsColumns.length; i++) {
      headerValues.push(preStart + "-" + preEnd, postStart + "-" + postEnd, "Change");
    }
   // headerValues.push("Scope");
    keywordSheet.appendRow(headerValues);
    keywordSheet.getRange("A1:CA1").setFontWeight("bold");
    
    var numColsToHide = 4 - segmentsToInclude.length;
    var segmentColIndex = 5 + segmentsToInclude.length;
    keywordSheet.hideColumns(segmentColIndex, numColsToHide);
    keywordSheet.setFrozenColumns(4);
  }
  
  if(includeAdLevel === true) {
    var adSheet = spreadsheet.insertSheet("Ads");
    
    adSheet.getRange("AA1:CA1").clear();
    for(var i = 0; i < metricsColumns.length; i++) {
      var startColNum = i*3 + 13;
      var endColNum = i*3 + 15;
      var startCol = alphabet[startColNum] + "1";
      var endCol = alphabet[endColNum] + "1";
      var rangeText = startCol + ":" + endCol;
      var metricName = metricsColumns[i];
      adSheet.getRange(rangeText).mergeAcross().setValue(metricName).setFontSize(10);
       adSheet.getRange("B3:BA1").sortOrder("ASCENDING");
    }
    //adjust and sort by ad group





    var startCol = alphabet[startColNum+3] + "1";
   // adSheet.getRange(startCol).setValue("Scope").setFontSize(8);
    
    adSheet.setFrozenRows(numFrozenRows);
    
    var headerValues = ["Account", "Campaign", "Ad Group", "Keyword", "Headline", "Line1", "Line2", "Vis URL", "Ad ID",
                         "Segment 1", "Segment 2", "Segment 3", "Segment 4"];
    for(var i = 0; i < metricsColumns.length; i++) {
      headerValues.push(preStart + "-" + preEnd, postStart + "-" + postEnd, "Change");
    }
   // headerValues.push("Scope");
    adSheet.appendRow(headerValues);
    adSheet.getRange("A1:CA1").setFontWeight("bold");
    var numColsToHide = 4 - segmentsToInclude.length;
    var segmentColIndex = 10 + segmentsToInclude.length;
    adSheet.hideColumns(segmentColIndex, numColsToHide);
  }
  
  
  
  
  
  
  // ACCOUNT 
    if(includeAccountLevel === true) {
    var elements = new Array();
    var segmentString = "";
    for(var count = 0; count < segmentsToInclude.length; count++) {
      var segment = segmentsToInclude[count];
      segmentString = segmentString + ", " + segment;
    }
    if(segmentString != ", ") {
        var query = 'SELECT ' + accountAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   ACCOUNT_PERFORMANCE_REPORT ' +
        'DURING ' + preStart + ',' + preEnd;
      if(DEBUG) Logger.log("query: " + query);
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
        var query = 'SELECT ' + accountAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   ACCOUNT_PERFORMANCE_REPORT ' +
        'DURING ' + preStart + ',' + preEnd
        if(DEBUG) Logger.log("query: " + query);
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } 
    var preRows = preReport.rows();
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + accountAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   ACCOUNT_PERFORMANCE_REPORT ' +
        'DURING ' + postStart + ',' + postEnd
        if(DEBUG) Logger.log("query: " + query);
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + accountAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   ACCOUNT_PERFORMANCE_REPORT ' +
        'DURING ' + postStart + ',' + postEnd
        if(DEBUG) Logger.log("query: " + query);
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    }
    var postRows = postReport.rows();
    
    
    var accountElements = processReports(preRows, postRows, numFrozenRows, segmentsToInclude, "Account", decimalPoint);
      accountAlertCount = LoopThroughElements(accountElements, segmentsToInclude, acctSheet, "Account");
      if(VERBOSE) Logger.log(accountAlertCount + " account alerts.");
      if(accountAlertCount) {
        emailBody += "<li>" + accountAlertCount + " account alerts.</li>";
      } else {
        spreadsheet.deleteSheet(acctSheet);
      }
  }
  
  // CAMPAIGN REPORTS
  if(includeCampaignLevel === true) {
    var elements = new Array();
    var segmentString = "";
    for(var count = 0; count < segmentsToInclude.length; count++) {
      var segment = segmentsToInclude[count];
      segmentString = segmentString + ", " + segment;
    }
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + campaignAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + campaignAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } 
    var preRows = preReport.rows();
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + campaignAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + campaignAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    }
    var postRows = postReport.rows();
    
  
    var campaignElements = processReports(preRows, postRows, numFrozenRows, segmentsToInclude, "Campaign", decimalPoint);
    campaignAlertCount = LoopThroughElements(campaignElements, segmentsToInclude, campaignSheet, "Campaign");
    if(VERBOSE) Logger.log(campaignAlertCount + " campaign alerts.");
    if(campaignAlertCount) {
      emailBody += "<li>" + campaignAlertCount + " campaign alerts</li>";
    } else {
      spreadsheet.deleteSheet(campaignSheet);
    }
  }
  
  // Ad GROUP REPORTS
  if(includeAdGroupLevel === true) {
    var elements = new Array();
    var segmentString = "";
    for(var count = 0; count < segmentsToInclude.length; count++) {
      var segment = segmentsToInclude[count];
      segmentString = segmentString + ", " + segment;
    }
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + adGroupAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   ADGROUP_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd;
      if(DEBUG) Logger.log("Query: " + query);
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});/*removed v201802 JAR*/
        
    } else {
      var query = 'SELECT ' + adGroupAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   ADGROUP_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd;
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } 
    var preRows = preReport.rows();
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + adGroupAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   ADGROUP_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + adGroupAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   ADGROUP_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    }
    var postRows = postReport.rows();
    
    
    var adGroupElements = processReports(preRows, postRows, numFrozenRows, segmentsToInclude, "Ad Group", decimalPoint);
    adGroupAlertCount = LoopThroughElements(adGroupElements, segmentsToInclude, adGroupSheet, "Ad Group");
    if(VERBOSE) Logger.log(adGroupAlertCount + " ad group alerts");
    if(adGroupAlertCount) {
      emailBody += "<li>" + adGroupAlertCount + " ad group alerts</li>";
    } else {
     spreadsheet.deleteSheet(adGroupSheet); 
    }
  }
  
  // KEYWORD REPORTS
  if(includeKeywordLevel === true) {
    var elements = new Array();
    var segmentString = "";
    for(var count = 0; count < segmentsToInclude.length; count++) {
      var segment = segmentsToInclude[count];
      segmentString = segmentString + ", " + segment;
    }
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + keywordAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   KEYWORDS_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + keywordAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   KEYWORDS_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd;
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } 
    var preRows = preReport.rows();
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + keywordAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   KEYWORDS_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + keywordAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   KEYWORDS_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    }
    var postRows = postReport.rows();
    
    
    var keywordElements = processReports(preRows, postRows, numFrozenRows, segmentsToInclude, "Keyword", decimalPoint);
    keywordAlertCount = LoopThroughElements(keywordElements, segmentsToInclude, keywordSheet, "Keyword");
    if(VERBOSE) Logger.log(keywordAlertCount + " keyword alerts.");
    if(keywordAlertCount) {
      emailBody += "<li>" + keywordAlertCount + " keyword alerts</li>";
    } else {
      spreadsheet.deleteSheet(keywordSheet);
    }
  }
  
  // AD REPORTS
  if(includeAdLevel === true) {
    var elements = new Array();
    var segmentString = "";
    for(var count = 0; count < segmentsToInclude.length; count++) {
      var segment = segmentsToInclude[count];
      segmentString = segmentString + ", " + segment;
    }
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + adAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   AD_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd;
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + adAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   AD_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + preStart + ',' + preEnd;
      var preReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } 
    var preRows = preReport.rows();
    
    if(segmentString != ", ") {
      var query = 'SELECT ' + adAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' + segmentString + " " +
        'FROM   AD_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    } else {
      var query = 'SELECT ' + adAttributeColumns.join(',') + "," + metricsColumns.join(',') + ' ' +
        'FROM   AD_PERFORMANCE_REPORT ' +
        'WHERE  Impressions > 0 ' +
        'AND ' + campaignNameSelectorStatement + ' ' +
        'DURING ' + postStart + ',' + postEnd;
      var postReport = AdWordsApp.report(query,{apiVersion:"v201802"});
    }
    var postRows = postReport.rows();
    
    
    var adElements = processReports(preRows, postRows, numFrozenRows, segmentsToInclude, "Ad", decimalPoint);
    adAlertCount = LoopThroughElements(adElements, segmentsToInclude, adSheet, "Ad");
    if(VERBOSE) Logger.log(adAlertCount + " ad alerts.");
    if(adAlertCount) {
      emailBody += "<li>" + adAlertCount + " ad alerts</li>";
    } else {
      spreadsheet.deleteSheet(adSheet);
    }
  
  }
  
  var totalAlertCount = accountAlertCount + campaignAlertCount + adGroupAlertCount + keywordAlertCount + adAlertCount;
  
  emailBody += "</ul>";
  emailBody += "Full list of alerts: " + spreadsheetUrl;
  
  if(totalAlertCount) {
    
    var emailType = "warning";
    var subject = totalAlertCount + " AdWords Alerts";
    
    sendEmailNotifications (currentSetting.emailAddresses, subject, emailBody, emailType );
    return "Ran successfully. Anomaly list: <a href=\""+spreadsheetUrl+"\" target=\"_blank\">link</a>";
  } else {
    var spreadsheetId = spreadsheet.getId();
    DriveApp.getFileById(spreadsheetId).setTrashed(true);
    return "Ran successfully. No anomalies found.";
  }
  
  
  
  
  
  
  
  
  
  
  
  function processReports(preRows, postRows, numFrozenRows, segmentsToInclude, reportType, decimalPoint) {
  var lineCounter = numFrozenRows;
    var kpiValues = new Array();
    var keys = new Array();
    
    
  while (preRows.hasNext()) {
    lineCounter++;
    //Logger.log(lineCounter);
    var row = preRows.next();

    var allConversions = parseInt(row['AllConversions']); 
    var allConversionValue = getFloat(row['AllConversionValue']);
    var allConversionRate = getFloat(row['AllConversionRate']);
    var avgCpc = getFloat(row['AverageCpc']);
    var ctr = getFloat(row['Ctr']);
    var impressions = parseInt(row['Impressions']);
    var clicks = parseInt(row['Clicks']);
    var avgPos = getFloat(row['AveragePosition']);
    var cost = row['Cost'];
    if(decimalPoint == ".") {
      cost = getFloat(cost.replace(/[^0-9-.]/g, ''));
    } else {
      cost = getFloat(cost.replace(/[^0-9-,]/g, ''));
    }
    var conversionRate = getFloat(row['ConversionRate']);
    var conversions = getFloat(row['Conversions']);
    var conversionValue = getFloat(row['ConversionValue']);
   
    //var convertedClicks = parseInt(row['ConvertedClicks']);
    var costPerConversion = row['CostPerConversion'];
    if(decimalPoint == ".") {
      costPerConversion = getFloat(costPerConversion.replace(/[^0-9-.]/g, ''));
    } else {
      costPerConversion = getFloat(costPerConversion.replace(/[^0-9-,]/g, ''));
    }

    var costPerAllConversion = row['CostPerAllConversion'];
    if(decimalPoint == ".") {
      costPerAllConversion = getFloat(costPerAllConversion.replace(/[^0-9-.]/g, ''));
    } else {
      costPerAllConversion = getFloat(costPerAllConversion.replace(/[^0-9-,]/g, ''));
    }
    var crossDeviceConversions = getFloat(row['CrossDeviceConversions']);
    var valuePerConversion = row['ValuePerConversion'];
    if(decimalPoint == ".") {
      valuePerConversion = getFloat(valuePerConversion.replace(/[^0-9-.]/g, ''));
    } else {
      valuePerConversion = getFloat(valuePerConversion.replace(/[^0-9-,]/g, ''));
    }
    var valuePerAllConversion = row['ValuePerAllConversion'];
    if(decimalPoint == ".") {
      valuePerAllConversion = getFloat(valuePerAllConversion.replace(/[^0-9-.]/g, ''));
    } else {
      valuePerAllConversion = getFloat(valuePerAllConversion.replace(/[^0-9-,]/g, ''));
    }
    var viewThroughConversions = getFloat(row['ViewThroughConversions']);
    
    if(reportType == "Account") {
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      var campaignName = "";
      var adGroupName = "";
      var keyword = "";
      var name = accountName;
    } else if(reportType == "Campaign") {
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      var campaignName = row['CampaignName'];
      var adGroupName = "";
      var keyword = "";
      var name = campaignName;
    } else if(reportType == "Ad Group") {
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      var keyword = "";
      var name = campaignName + " - " + adGroupName;
    } else if(reportType == "Keyword") {
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      var keyword = row['Criteria'];
      var name = campaignName + " - " + adGroupName + " - " + keyword;
    } else if(reportType == "Ad") {
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      var campaignName = row['CampaignName'];
      var adGroupName = row['AdGroupName'];
      var keyword = "";
      var headline = row['Headline'];
      var line1 = row['Description1'];
      var line2 = row['Description2'];
      var visibleUrl = row['DisplayUrl'];
      var adId = row['Id'];
      var name = adId;
    }
    
    
    //var segmentValues = new Array();
    for(var segmentCount = 0; segmentCount < segmentsToInclude.length; segmentCount++) {
      var key = segmentsToInclude[segmentCount];
      var value = row[key];
      if(segmentCount == 0) var segment1 = value;
      if(segmentCount == 1) var segment2 = value;
      if(segmentCount == 2) var segment3 = value;
      if(segmentCount == 3) var segment4 = value;
      if(segmentCount == 4) var segment5 = value;
      if(segmentCount == 5) var segment6 = value;      
    }
    
    if(segmentCount == 0) {
      if(!elements[name]) {
        elements[name] = {};
        elements[name].pre = {};
        elements[name].post = {};
      }
      
      
      elements[name].pre.AllConversions = allConversions;
      elements[name].pre.AllConversionValue = allConversionValue;
      elements[name].pre.AllConversionRate = allConversionRate;
      elements[name].pre.AverageCpc = avgCpc;
      elements[name].pre.Ctr = ctr;
      elements[name].pre.Impressions = impressions;
      elements[name].pre.Clicks = clicks;
      elements[name].pre.AveragePosition = avgPos;
      elements[name].pre.Cost = cost;
      elements[name].pre.ConversionRate = conversionRate;
      elements[name].pre.Conversions = conversions;
      elements[name].pre.ConversionValue = conversionValue;
     // elements[name].pre.ConvertedClicks = convertedClicks;
      elements[name].pre.CostPerConversion = costPerConversion;
      elements[name].pre.CostPerAllConversion = costPerAllConversion;
      elements[name].pre.CrossDeviceConversions = crossDeviceConversions;
      elements[name].pre.ValuePerConversion = valuePerConversion;
      elements[name].pre.ValuePerAllConversion = valuePerAllConversion;
      elements[name].pre.ViewThroughConversions = viewThroughConversions;
      
      elements[name].post.AllConversions = 0;
      elements[name].post.AllConversionValue = 0;
      elements[name].post.AllConversionRate = 0;
      elements[name].post.AverageCpc = 0;
      elements[name].post.Ctr = 0;
      elements[name].post.Impressions = 0;
      elements[name].post.Clicks = 0;
      elements[name].post.AveragePosition = 0;
      elements[name].post.Cost = 0;
      elements[name].post.ConversionRate = 0;
      elements[name].post.Conversions = 0;
      elements[name].post.ConversionValue = 0;
    //  elements[name].post.ConvertedClicks = 0;
      elements[name].post.CostPerConversion = 0;
      elements[name].post.CostPerAllConversion = 0;
      elements[name].post.CrossDeviceConversions = 0;
      elements[name].post.ValuePerConversion = 0;
      elements[name].post.ValuePerAllConversion = 0;
      elements[name].post.ViewThroughConversions = 0;
      
      elements[name].campaignName = campaignName;
      elements[name].adGroupName = adGroupName;
      elements[name].accountName = accountName;
      elements[name].keyword = keyword;
      elements[name].headline = headline;
      elements[name].line1 = line1;
      elements[name].line2 = line2;
      elements[name].visibleUrl = visibleUrl;
      elements[name].adId = adId;
      keys.push(name);
      kpiValues.push(conversions);
    } else if(segmentCount == 1) {
      if(!elements[name]) {
        elements[name] = {};
      }
      if(!elements[name][segment1]) {
        elements[name][segment1] = {};
        elements[name][segment1].pre = {};
        elements[name][segment1].post = {};
      }
      elements[name][segment1].pre.AllConversions = allConversions;
      elements[name][segment1].pre.AllConversionValue = allConversionValue;
      elements[name][segment1].pre.AllConversionRate = allConversionRate;
      elements[name][segment1].pre.AverageCpc = avgCpc;
      elements[name][segment1].pre.Ctr = ctr;
      elements[name][segment1].pre.Impressions = impressions;
      elements[name][segment1].pre.Clicks = clicks;
      elements[name][segment1].pre.AveragePosition = avgPos;
      elements[name][segment1].pre.Cost = cost;
      elements[name][segment1].pre.ConversionRate = conversionRate;
      elements[name][segment1].pre.Conversions = conversions;
      elements[name][segment1].pre.ConversionValue = conversionValue;
   //   elements[name][segment1].pre.ConvertedClicks = convertedClicks;
      elements[name][segment1].pre.CostPerConversion = costPerConversion;
      elements[name][segment1].pre.CostPerAllConversion = costPerAllConversion;
      elements[name][segment1].pre.CrossDeviceConversions = crossDeviceConversions;
      elements[name][segment1].pre.ValuePerConversion = valuePerConversion;
      elements[name][segment1].pre.ValuePerAllConversion = valuePerAllConversion;
      elements[name][segment1].pre.ViewThroughConversions = viewThroughConversions;
      
      elements[name][segment1].post.AllConversions = 0;
      elements[name][segment1].post.AllConversionValue = 0;
      elements[name][segment1].post.AllConversionRate = 0;
      elements[name][segment1].post.AverageCpc = 0;
      elements[name][segment1].post.Ctr = 0;
      elements[name][segment1].post.Impressions = 0;
      elements[name][segment1].post.Clicks = 0;
      elements[name][segment1].post.AveragePosition = 0;
      elements[name][segment1].post.Cost = 0;
      elements[name][segment1].post.ConversionRate = 0;
      elements[name][segment1].post.Conversions = 0;
      elements[name][segment1].post.ConversionValue = 0;
     // elements[name][segment1].post.ConvertedClicks = 0;
      elements[name][segment1].post.CostPerConversion = 0;
      elements[name][segment1].post.CostPerAllConversion = 0;
      elements[name][segment1].post.CrossDeviceConversions = 0;
      elements[name][segment1].post.ValuePerConversion = 0;
      elements[name][segment1].post.ValuePerAllConversion = 0;
      elements[name][segment1].post.ViewThroughConversions = 0;
      
      
      elements[name][segment1].campaignName = campaignName;
      elements[name][segment1].adGroupName = adGroupName;
      elements[name][segment1].accountName = accountName;
      elements[name][segment1].keyword = keyword;
      elements[name][segment1].headline = headline;
      elements[name][segment1].line1 = line1;
      elements[name][segment1].line2 = line2;
      elements[name][segment1].visibleUrl = visibleUrl;
      elements[name][segment1].adId = adId;
      var keyName = name + "../|\.." + segment1;
      keys.push(keyName);
      kpiValues.push(conversions);
    } else if(segmentCount == 2) {
      if(!elements[name]) {
        elements[name] = {};
      }
      if(!elements[name][segment1]) {
        elements[name][segment1] = {};
      }
      if(!elements[name][segment1][segment2]) {
        elements[name][segment1][segment2] = {};
        elements[name][segment1][segment2].pre = {};
        elements[name][segment1][segment2].post = {};
      }
      
      elements[name][segment1][segment2].pre.AllConversions = allConversions;
      /*elements[name][segment1][segment2].pre.AllConversionValue = allConversionValue;
      elements[name][segment1][segment2].pre.AllConversionRate = allConversionRate;
      elements[name][segment1][segment2].pre.AverageCpc = avgCpc;*/
      elements[name][segment1][segment2].pre.Ctr = ctr;
      elements[name][segment1][segment2].pre.Impressions = impressions;
      elements[name][segment1][segment2].pre.Clicks = clicks;
      elements[name][segment1][segment2].pre.AveragePosition = avgPos;
      elements[name][segment1][segment2].pre.Cost = cost;
      elements[name][segment1][segment2].pre.ConversionRate = conversionRate;
      elements[name][segment1][segment2].pre.Conversions = conversions;
      elements[name][segment1][segment2].pre.ConversionValue = conversionValue;
   //   elements[name][segment1][segment2].pre.ConvertedClicks = convertedClicks;
      elements[name][segment1][segment2].pre.CostPerConversion = costPerConversion;
      elements[name][segment1][segment2].pre.CostPerAllConversion = costPerAllConversion;
      elements[name][segment1][segment2].pre.CrossDeviceConversions = crossDeviceConversions;
      elements[name][segment1][segment2].pre.ValuePerConversion = valuePerConversion;
      elements[name][segment1][segment2].pre.ValuePerAllConversion = valuePerAllConversion;
      elements[name][segment1][segment2].pre.ViewThroughConversions = viewThroughConversions;
      
      elements[name][segment1][segment2].post.AllConversions = 0;
      elements[name][segment1][segment2].post.AllConversionValue = 0;
      elements[name][segment1][segment2].post.AllConversionRate = 0;
      elements[name][segment1][segment2].post.AverageCpc = 0;
      elements[name][segment1][segment2].post.Ctr = 0;
      elements[name][segment1][segment2].post.Impressions = 0;
      elements[name][segment1][segment2].post.Clicks = 0;
      elements[name][segment1][segment2].post.AveragePosition = 0;
      elements[name][segment1][segment2].post.Cost = 0;
      elements[name][segment1][segment2].post.ConversionRate = 0;
      elements[name][segment1][segment2].post.Conversions = 0;
      elements[name][segment1][segment2].post.ConversionValue = 0;
    //  elements[name][segment1][segment2].post.ConvertedClicks = 0;
      elements[name][segment1][segment2].post.CostPerConversion = 0;
      elements[name][segment1][segment2].post.CostPerAllConversion = 0;
      elements[name][segment1][segment2].post.CrossDeviceConversions = 0;
      elements[name][segment1][segment2].post.ValuePerConversion = 0;
      elements[name][segment1][segment2].post.ValuePerAllConversion = 0;
      elements[name][segment1][segment2].post.ViewThroughConversions = 0;
      
      elements[name][segment1][segment2].campaignName = campaignName;
      elements[name][segment1][segment2].adGroupName = adGroupName;
      elements[name][segment1][segment2].accountName = accountName;
      elements[name][segment1][segment2].keyword = keyword;
      elements[name][segment1][segment2].headline = headline;
      elements[name][segment1][segment2].line1 = line1;
      elements[name][segment1][segment2].line2 = line2;
      elements[name][segment1][segment2].visibleUrl = visibleUrl;
      elements[name][segment1][segment2].adId = adId;
      var keyName = name + "../|\.." + segment1 + "../|\.." + segment2;
      keys.push(keyName);
      kpiValues.push(conversions);
    }
    
     
  }
  
    while(postRows.hasNext()) {
      var row = postRows.next();
      var allConversions = parseInt(row['AllConversions']); 
    var allConversionValue = getFloat(row['AllConversionValue']);
    var allConversionRate = getFloat(row['AllConversionRate']);
    var avgCpc = getFloat(row['AverageCpc']);
    var ctr = getFloat(row['Ctr']);
    var impressions = parseInt(row['Impressions']);
    var clicks = parseInt(row['Clicks']);
    var avgPos = getFloat(row['AveragePosition']);
    var cost = row['Cost'];
    if(decimalPoint == ".") {
      cost = getFloat(cost.replace(/[^0-9-.]/g, ''));
    } else {
      cost = getFloat(cost.replace(/[^0-9-,]/g, ''));
    }
    var conversionRate = getFloat(row['ConversionRate']);
    var conversions = getFloat(row['Conversions']);
    var conversionValue = getFloat(row['ConversionValue']);
   // var convertedClicks = parseInt(row['ConvertedClicks']);
    var costPerConversion = row['CostPerConversion'];
    if(decimalPoint == ".") {
      costPerConversion = getFloat(costPerConversion.replace(/[^0-9-.]/g, ''));
    } else {
      costPerConversion = getFloat(costPerConversion.replace(/[^0-9-,]/g, ''));
    }
    var costPerAllConversion = row['CostPerAllConversion'];
    if(decimalPoint == ".") {
      costPerAllConversion = getFloat(costPerAllConversion.replace(/[^0-9-.]/g, ''));
    } else {
      costPerAllConversion = getFloat(costPerAllConversion.replace(/[^0-9-,]/g, ''));
    }
    var crossDeviceConversions = getFloat(row['CrossDeviceConversions']);
    var valuePerConversion = row['ValuePerConversion'];
    if(decimalPoint == ".") {
      valuePerConversion = getFloat(valuePerConversion.replace(/[^0-9-.]/g, ''));
    } else {
      valuePerConversion = getFloat(valuePerConversion.replace(/[^0-9-,]/g, ''));
    }
    var valuePerAllConversion = row['ValuePerAllConversion'];
    if(decimalPoint == ".") {
      valuePerAllConversion = getFloat(valuePerAllConversion.replace(/[^0-9-.]/g, ''));
    } else {
      valuePerAllConversion = getFloat(valuePerAllConversion.replace(/[^0-9-,]/g, ''));
    }
    var viewThroughConversions = getFloat(row['ViewThroughConversions']);
      
      
      var accountName = row['AccountDescriptiveName'];
      var accountId = row['AccountId'];
      if(reportType == "Account") {
        var accountName = row['AccountDescriptiveName'];
        var accountId = row['AccountId'];
        var name = accountName;
        var campaignName = "";
        var adGroupName = "";
        var keyword = "";
      } else if(reportType == "Campaign") {
        var accountName = row['AccountDescriptiveName'];
        var accountId = row['AccountId'];
        var campaignName = row['CampaignName'];
        var adGroupName = "";
        var keyword = "";
        var name = campaignName;
      } else if(reportType == "Ad Group") {
        var accountName = row['AccountDescriptiveName'];
        var accountId = row['AccountId'];
        var campaignName = row['CampaignName'];
        var adGroupName = row['AdGroupName'];
        var keyword = "";
        var name = campaignName + " - " + adGroupName;
      } else if(reportType == "Keyword") {
        var accountName = row['AccountDescriptiveName'];
        var accountId = row['AccountId'];
        var campaignName = row['CampaignName'];
        var adGroupName = row['AdGroupName'];
        var keyword = row['Criteria'];
        var name = campaignName + " - " + adGroupName + " - " + keyword;
      } else if(reportType == "Ad") {
        var accountName = row['AccountDescriptiveName'];
        var accountId = row['AccountId'];
        var campaignName = row['CampaignName'];
        var adGroupName = row['AdGroupName'];
        var keyword = "";
        var headline = row['Headline'];
        var line1 = row['Description1'];
        var line2 = row['Description2'];
        var visibleUrl = row['DisplayUrl'];
        var adId = row['Id'];
        var name = adId;
      }


      for(var segmentCount = 0; segmentCount < segmentsToInclude.length; segmentCount++) {
        var key = segmentsToInclude[segmentCount];
      var value = row[key];
      if(segmentCount == 0) var segment1 = value;
      if(segmentCount == 1) var segment2 = value;
      if(segmentCount == 2) var segment3 = value;
      if(segmentCount == 3) var segment4 = value;
      if(segmentCount == 4) var segment5 = value;
      if(segmentCount == 5) var segment6 = value;
    }
    
    if(segmentCount == 0) {
      if(!elements[name]) {
        elements[name] = {};
        elements[name].pre = {};
        elements[name].post = {};
      } 
      elements[name].post.AllConversions = allConversions;
      elements[name].post.AllConversionValue = allConversionValue;
      elements[name].post.AllConversionRate = allConversionRate;
      elements[name].post.AverageCpc = avgCpc;
      elements[name].post.Ctr = ctr;
      elements[name].post.Impressions = impressions;
      elements[name].post.Clicks = clicks;
      elements[name].post.AveragePosition = avgPos;
      elements[name].post.Cost = cost;
      elements[name].post.ConversionRate = conversionRate;
      elements[name].post.Conversions = conversions;
      elements[name].post.ConversionValue = conversionValue;
    //  elements[name].post.ConvertedClicks = convertedClicks;
      elements[name].post.CostPerConversion = costPerConversion;
      elements[name].post.CostPerAllConversion = costPerAllConversion;
      elements[name].post.CrossDeviceConversions = crossDeviceConversions;
      elements[name].post.ValuePerConversion = valuePerConversion;
      elements[name].post.ValuePerAllConversion = valuePerAllConversion;
      elements[name].post.ViewThroughConversions = viewThroughConversions;
      
      elements[name].campaignName = campaignName;
      elements[name].adGroupName = adGroupName;
      elements[name].accountName = accountName;
      elements[name].keyword = keyword;
      elements[name].headline = headline;
      elements[name].line1 = line1;
      elements[name].line2 = line2;
      elements[name].visibleUrl = visibleUrl;
      elements[name].adId = adId;
    
      
      if(!elements[name].pre.Impressions) {
        elements[name].pre.AllConversions = 0;
        elements[name].pre.AllConversionValue = 0;
        elements[name].pre.AllConversionRate = 0;
        elements[name].pre.AverageCpc = 0;
        elements[name].pre.Ctr = 0;
        elements[name].pre.Impressions = 0;
        elements[name].pre.Clicks = 0;
        elements[name].pre.AveragePosition = 0;
        elements[name].pre.Cost = 0;
        elements[name].pre.ConversionRate = 0;
        elements[name].pre.Conversions = 0;
        elements[name].pre.ConversionValue = 0;
     //   elements[name].pre.ConvertedClicks = 0;
        elements[name].pre.CostPerConversion = 0;
        elements[name].pre.CostPerAllConversion = 0;
        elements[name].pre.CrossDeviceConversions = 0;
        elements[name].pre.ValuePerConversion = 0;
        elements[name].pre.ValuePerAllConversion = 0;
        elements[name].pre.ViewThroughConversions = 0;
      }
      
    } else if(segmentCount == 1) {
      if(!elements[name]) {
        elements[name] = {};
      }
      if(!elements[name][segment1]) {
        elements[name][segment1] = {};
        elements[name][segment1].pre = {};
        elements[name][segment1].post = {};
      }
      elements[name][segment1].post.AllConversions = allConversions;
      elements[name][segment1].post.AllConversionValue = allConversionValue;
      elements[name][segment1].post.AllConversionRate = allConversionRate;
      elements[name][segment1].post.AverageCpc = avgCpc;
      elements[name][segment1].post.Ctr = ctr;
      elements[name][segment1].post.Impressions = impressions;
      elements[name][segment1].post.Clicks = clicks;
      elements[name][segment1].post.AveragePosition = avgPos;
      elements[name][segment1].post.Cost = cost;
      elements[name][segment1].post.ConversionRate = conversionRate;
      elements[name][segment1].post.Conversions = conversions;
      elements[name][segment1].post.ConversionValue = conversionValue;
   //   elements[name][segment1].post.ConvertedClicks = convertedClicks;
      elements[name][segment1].post.CostPerConversion = costPerConversion;
      elements[name][segment1].post.CostPerAllConversion = costPerAllConversion;
      elements[name][segment1].post.CrossDeviceConversions = crossDeviceConversions;
      elements[name][segment1].post.ValuePerConversion = valuePerConversion;
      elements[name][segment1].post.ValuePerAllConversion = valuePerAllConversion;
      elements[name][segment1].post.ViewThroughConversions = viewThroughConversions;
      
      elements[name][segment1].campaignName = campaignName;
      elements[name][segment1].adGroupName = adGroupName;
      elements[name][segment1].accountName = accountName;
      elements[name][segment1].keyword = keyword;
      elements[name][segment1].headline = headline;
      elements[name][segment1].line1 = line1;
      elements[name][segment1].line2 = line2;
      elements[name][segment1].visibleUrl = visibleUrl;
      elements[name][segment1].adId = adId;
    
    
      if(elements[name][segment1].pre.Impressions == "undefined") {
        elements[name][segment1].pre.AllConversions = 0;
        elements[name][segment1].pre.AllConversionValue = 0;
        elements[name][segment1].pre.AllConversionRate = 0;
        elements[name][segment1].pre.AverageCpc = 0;
        elements[name][segment1].pre.Ctr = 0;
        elements[name][segment1].pre.Impressions = 0;
        elements[name][segment1].pre.Clicks = 0;
        elements[name][segment1].pre.AveragePosition = 0;
        elements[name][segment1].pre.Cost = 0;
        elements[name][segment1].pre.ConversionRate = 0;
        elements[name][segment1].pre.Conversions = 0;
        elements[name][segment1].pre.ConversionValue = 0;
   //    elements[name][segment1].pre.ConvertedClicks = 0;
        elements[name][segment1].pre.CostPerConversion = 0;
        elements[name][segment1].pre.CostPerAllConversion = 0;
        elements[name][segment1].pre.CrossDeviceConversions = 0;
        elements[name][segment1].pre.ValuePerConversion = 0;
        elements[name][segment1].pre.ValuePerAllConversion = 0;
        elements[name][segment1].pre.ViewThroughConversions = 0;
      }
      
    } else if(segmentCount == 2) {
      if(!elements[name]) {
        elements[name] = {};
      }
      if(!elements[name][segment1]) {
        elements[name][segment1] = {};
      }
      if(!elements[name][segment1][segment2]) {
        elements[name][segment1][segment2] = {};
        elements[name][segment1][segment2].pre = {};
        elements[name][segment1][segment2].post = {};
      }
      elements[name][segment1][segment2].post.AllConversions = allConversions;
      elements[name][segment1][segment2].post.AllConversionValue = allConversionValue;
      elements[name][segment1][segment2].post.AllConversionRate = allConversionRate;
      elements[name][segment1][segment2].post.AverageCpc = avgCpc;
      elements[name][segment1][segment2].post.Ctr = ctr;
      elements[name][segment1][segment2].post.Impressions = impressions;
      elements[name][segment1][segment2].post.Clicks = clicks;
      elements[name][segment1][segment2].post.AveragePosition = avgPos;
      elements[name][segment1][segment2].post.Cost = cost;
      elements[name][segment1][segment2].post.ConversionRate = conversionRate;
      elements[name][segment1][segment2].post.Conversions = conversions;
      elements[name][segment1][segment2].post.ConversionValue = conversionValue;
    //  elements[name][segment1][segment2].post.ConvertedClicks = convertedClicks;
      elements[name][segment1][segment2].post.CostPerConversion = costPerConversion;
      elements[name][segment1][segment2].post.CostPerAllConversion = costPerAllConversion;
      elements[name][segment1][segment2].post.CrossDeviceConversions = crossDeviceConversions;
      elements[name][segment1][segment2].post.ValuePerConversion = valuePerConversion;
      elements[name][segment1][segment2].post.ValuePerAllConversion = valuePerAllConversion;
      elements[name][segment1][segment2].post.ViewThroughConversions = viewThroughConversions;
      
      elements[name][segment1][segment2].campaignName = campaignName;
      elements[name][segment1][segment2].adGroupName = adGroupName;
      elements[name][segment1][segment2].accountName = accountName;
      elements[name][segment1][segment2].keyword = keyword;
      elements[name][segment1][segment2].headline = headline;
      elements[name][segment1][segment2].line1 = line1;
      elements[name][segment1][segment2].line2 = line2;
      elements[name][segment1][segment2].visibleUrl = visibleUrl;
      elements[name][segment1][segment2].adId = adId;
    
      
      if(elements[name][segment1][segment2].pre.Impressions == "undefined") {
        elements[name][segment1][segment2].pre.AllConversions = 0;
        elements[name][segment1][segment2].pre.AllConversionValue = 0;
        elements[name][segment1][segment2].pre.AllConversionRate = 0;
        elements[name][segment1][segment2].pre.AverageCpc = 0;
        elements[name][segment1][segment2].pre.Ctr = 0;
        elements[name][segment1][segment2].pre.Impressions = 0;
        elements[name][segment1][segment2].pre.Clicks = 0;
        elements[name][segment1][segment2].pre.AveragePosition = 0;
        elements[name][segment1][segment2].pre.Cost = 0;
        elements[name][segment1][segment2].pre.ConversionRate = 0;
        elements[name][segment1][segment2].pre.Conversions = 0;
        elements[name][segment1][segment2].pre.ConversionValue = 0;
     //   elements[name][segment1][segment2].pre.ConvertedClicks = 0;
        elements[name][segment1][segment2].pre.CostPerConversion = 0;
        elements[name][segment1][segment2].pre.CostPerAllConversion = 0;
        elements[name][segment1][segment2].pre.CrossDeviceConversions = 0;
        elements[name][segment1][segment2].pre.ValuePerConversion = 0;
        elements[name][segment1][segment2].pre.ValuePerAllConversion = 0;
        elements[name][segment1][segment2].pre.ViewThroughConversions = 0;
      }
    }
    
    
    
  } // while(postRows.hasNext()) 
    
    
    // do some sorting
    var list = [];
    for (var j=0; j<keys.length; j++) {
      list.push({'name': keys[j], 'value': kpiValues[j]});
    }
    
    list.sort(function(a, b) {
      return b.value - a.value;
    });
    
    for (var k=0; k<list.length; k++) {
      keys[k] = list[k].name;
      kpiValues[k] = list[k].value;
    }
    
    
    elements.keys = keys;
    elements.kpiValues = kpiValues;
    return elements;
  }
  




  
  
  
  
  
  
  
  
}

function LoopThroughElements(elements, segmentsToInclude, optSheet, reportType) {
    
    var keys = elements.keys;
    var kpiValues = elements.kpiValues;
  var alertCounter = 0;
    
    for(var elementCount = 0; elementCount < keys.length; elementCount++) {
      var element = keys[elementCount];
      var elementParts = element.split("../|\..");
      var name = elementParts[0];
      var segment1 = elementParts[1];
      var segment2 = elementParts[2];
      var segment3 = elementParts[3];
      
      if(DEBUG) Logger.log("element unique ID: " + name);
      var elementOut = elements[name];
      if(segmentsToInclude.length == 0) {
        var alerted = analyze(elementOut, optSheet, reportType, element);
        if(alerted) alertCounter++;
      }
      if(segmentsToInclude.length >= 1) {
        if(DEBUG) Logger.log(segment1);
        var elementOut = elements[name][segment1];
        if(segmentsToInclude.length == 1) {
          var alerted = analyze(elementOut, optSheet, reportType, element, segment1);
          if(alerted) alertCounter++;
        }
        if(segmentsToInclude.length >= 2) {
          if(DEBUG) Logger.log(segment2);
          var elementOut = elements[name][segment1][segment2];
          if(segmentsToInclude.length == 2) {
            var alerted = analyze(elementOut, optSheet, reportType, element, segment1, segment2);
            if(alerted) alertCounter++;
          }
          if(segmentsToInclude.length >= 3) {
            if(DEBUG) Logger.log(segment3);
            var elementOut = elements[name][segment1][segment2][segment3];
            if(segmentsToInclude.length == 3) {
              var alerted = analyze(elementOut, optSheet, reportType, element, segment1, segment2, segment3);
              if(alerted) alertCounter++;
            }
          } 
        } 
      }
    }
  return(alertCounter);
  }

function analyze(element, sheet, scope, name, segment1, segment2, segment3, segment4) {
    
    var rows = sheet.getDataRange();
    var numRows = rows.getNumRows();
    var numCols = rows.getLastColumn();
    var lineCounter = numRows + 1;
  
  
  var alert = false;
  var goodTextColor = "green";
  var badTextColor = "red";
  var goodCellColor = "#d9ffcc";
  var badCellColor = "#ffcccc";
   
  
  var preAllConversions = element.pre.AllConversions;
    var postAllConversions = element.post.AllConversions;
    if(preAllConversions > 0) {
      var AllConversionsChange = postAllConversions - preAllConversions;
      var AllConversionsChangePercent = (postAllConversions - preAllConversions) / preAllConversions;
      var AllConversionsTextColor = (AllConversionsChange > 0 ? currentSetting.colors["AllConversions"].textColorForIncreases : currentSetting.colors["AllConversions"].textColorForDecreases);
    } else if (preAllConversions == 0 && postAllConversions > 0) {
      var AllConversionsChange = postAllConversions;
      var AllConversionsChangePercent = "\u221E";
      var AllConversionsColor = currentSetting.colors["AllConversions"].textColorForIncreases;
    } else if (preAllConversions == postAllConversions) {
      var AllConversionsChange = 0;
      var AllConversionsChangePercent = 0;
    }
  var AllConversionsCellColor = "white";
  if(preAllConversions >= currentSetting.minAlertAllConversions || postAllConversions >= currentSetting.minAlertAllConversions) {
    if(currentSetting.minIncreaseForAllConversionsAlert && AllConversionsChangePercent > 0 && AllConversionsChangePercent > currentSetting.minIncreaseForAllConversionsAlert) {
      var alert = true;
      AllConversionsCellColor = currentSetting.colors["AllConversions"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForAllConversionsAlert && AllConversionsChangePercent < 0 && AllConversionsChangePercent < currentSetting.minDecreaseForAllConversionsAlert) {
      var alert = true;
      AllConversionsCellColor = currentSetting.colors["AllConversions"].bgColorForDecreases;
    } 
  }
  
  var preAllConversionValue = element.pre.AllConversionValue;
    var postAllConversionValue = element.post.AllConversionValue;
    if(preAllConversionValue > 0) {
      var AllConversionValueChange = postAllConversionValue - preAllConversionValue;
      var AllConversionValueChangePercent = (postAllConversionValue - preAllConversionValue) / preAllConversionValue;
      var AllConversionValueTextColor = (AllConversionValueChange > 0 ? currentSetting.colors["AllConversionValue"].textColorForIncreases : currentSetting.colors["AllConversionValue"].textColorForDecreases);
    } else if (preAllConversionValue == 0 && postAllConversionValue > 0) {
      var AllConversionValueChange = postAllConversionValue;
      var AllConversionValueChangePercent = "\u221E";
      var AllConversionValueColor = currentSetting.colors["AllConversionValue"].textColorForIncreases;
    } else if (preAllConversionValue == postAllConversionValue) {
      var AllConversionValueChange = 0;
      var AllConversionValueChangePercent = 0;
    }
  var AllConversionValueCellColor = "white";
  if(preAllConversionValue >= currentSetting.minAlertAllConversionValue || postAllConversionValue >= currentSetting.minAlertAllConversionValue) {
    if(currentSetting.minIncreaseForAllConversionValueAlert && AllConversionValueChangePercent > 0 && AllConversionValueChangePercent > currentSetting.minIncreaseForAllConversionValueAlert) {
      var alert = true;
      AllConversionValueCellColor = currentSetting.colors["AllConversionValue"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForAllConversionValueAlert && AllConversionValueChangePercent < 0 && AllConversionValueChangePercent < currentSetting.minDecreaseForAllConversionValueAlert) {
      var alert = true;
      AllConversionValueCellColor = currentSetting.colors["AllConversionValue"].bgColorForDecreases;
    } 
  }
  
  
  var preAllConversionRate = element.pre.AllConversionRate;
  var postAllConversionRate = element.post.AllConversionRate;
  if(DEBUG) Logger.log("preAllConversionRate: " + preAllConversionRate);
  if(DEBUG) Logger.log("postAllConversionRate: " + postAllConversionRate);
    if(preAllConversionRate > 0) {
      var AllConversionRateChange = postAllConversionRate - preAllConversionRate;
      var AllConversionRateChangePercent = (postAllConversionRate - preAllConversionRate) / preAllConversionRate;
      var AllConversionRateTextColor = (AllConversionRateChange > 0 ? currentSetting.colors["AllConversionRate"].textColorForIncreases : currentSetting.colors["AllConversionRate"].textColorForDecreases);
    } else if (preAllConversionRate == 0 && postAllConversionRate > 0) {
      var AllConversionRateChange = postAllConversionRate;
      var AllConversionRateChangePercent = "\u221E";
      var AllConversionRateColor = currentSetting.colors["AllConversionRate"].textColorForIncreases;
    } else if (preAllConversionRate == postAllConversionRate) {
      var AllConversionRateChange = 0;
      var AllConversionRateChangePercent = 0;
    }
  var AllConversionRateCellColor = "white";
  if(preAllConversionRate >= currentSetting.minAlertAllConversionRate || postAllConversionRate >= currentSetting.minAlertAllConversionRate) {
    if(currentSetting.minIncreaseForAllConversionRateAlert && AllConversionRateChangePercent > 0 && AllConversionRateChangePercent > currentSetting.minIncreaseForAllConversionRateAlert) {
      var alert = true;
      AllConversionRateCellColor = currentSetting.colors["AllConversionRate"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForAllConversionRateAlert && AllConversionRateChangePercent < 0 && AllConversionRateChangePercent < currentSetting.minDecreaseForAllConversionRateAlert) {
      var alert = true;
      AllConversionRateCellColor = currentSetting.colors["AllConversionRate"].bgColorForDecreases;
    } 
  }
  
  
  var preAverageCpc = element.pre.AverageCpc;
    var postAverageCpc = element.post.AverageCpc;
    if(preAverageCpc > 0) {
      var AverageCpcChange = postAverageCpc - preAverageCpc;
      var AverageCpcChangePercent = (postAverageCpc - preAverageCpc) / preAverageCpc;
      var AverageCpcTextColor = (AverageCpcChange > 0 ? currentSetting.colors["AverageCpc"].textColorForIncreases : currentSetting.colors["AverageCpc"].textColorForDecreases);
    } else if (preAverageCpc == 0 && postAverageCpc > 0) {
      var AverageCpcChange = postAverageCpc;
      var AverageCpcChangePercent = "\u221E";
      var AverageCpcColor = currentSetting.colors["AverageCpc"].textColorForIncreases;
    } else if (preAverageCpc == postAverageCpc) {
      var AverageCpcChange = 0;
      var AverageCpcChangePercent = 0;
    }
  var AverageCpcCellColor = "white";
  if(preAverageCpc >= currentSetting.minAlertAverageCpc || postAverageCpc >= currentSetting.minAlertAverageCpc) {
    if(currentSetting.minIncreaseForAverageCpcAlert && AverageCpcChangePercent > 0 && AverageCpcChangePercent > currentSetting.minIncreaseForAverageCpcAlert) {
      var alert = true;
      AverageCpcCellColor = currentSetting.colors["AverageCpc"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForAverageCpcAlert && AverageCpcChangePercent < 0 && AverageCpcChangePercent < currentSetting.minDecreaseForAverageCpcAlert) {
      var alert = true;
      AverageCpcCellColor = currentSetting.colors["AverageCpc"].bgColorForDecreases;
    } 
  }
  
  var preCtr = element.pre.Ctr;
    var postCtr = element.post.Ctr;
    if(preCtr > 0) {
      var CtrChange = postCtr - preCtr;
      var CtrChangePercent = (postCtr - preCtr) / preCtr;
      var CtrTextColor = (CtrChange > 0 ? currentSetting.colors["Ctr"].textColorForIncreases : currentSetting.colors["Ctr"].textColorForDecreases);
    } else if (preCtr == 0 && postCtr > 0) {
      var CtrChange = postCtr;
      var CtrChangePercent = "\u221E";
      var CtrColor = currentSetting.colors["Ctr"].textColorForIncreases;
    } else if (preCtr == postCtr) {
      var CtrChange = 0;
      var CtrChangePercent = 0;
    }
  var CtrCellColor = "white";
  if(preCtr >= currentSetting.minAlertCtr || postCtr >= currentSetting.minAlertCtr) {
    if(currentSetting.minIncreaseForCtrAlert && CtrChangePercent > 0 && CtrChangePercent > currentSetting.minIncreaseForCtrAlert) {
      var alert = true;
      CtrCellColor = currentSetting.colors["Ctr"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForCtrAlert && CtrChangePercent < 0 && CtrChangePercent < currentSetting.minDecreaseForCtrAlert) {
      var alert = true;
      CtrCellColor = currentSetting.colors["Ctr"].bgColorForDecreases;
    } 
  }
  
  
  
  
  var preImpressions = element.pre.Impressions;
    var postImpressions = element.post.Impressions;
    if(preImpressions > 0) {
      var ImpressionsChange = postImpressions - preImpressions;
      var ImpressionsChangePercent = (postImpressions - preImpressions) / preImpressions;
      var ImpressionsTextColor = (ImpressionsChange > 0 ? currentSetting.colors["Impressions"].textColorForIncreases : currentSetting.colors["Impressions"].textColorForDecreases);
    } else if (preImpressions == 0 && postImpressions > 0) {
      var ImpressionsChange = postImpressions;
      var ImpressionsChangePercent = "\u221E";
      var ImpressionsColor = currentSetting.colors["Impressions"].textColorForIncreases;
    } else if (preImpressions == postImpressions) {
      var ImpressionsChange = 0;
      var ImpressionsChangePercent = 0;
    }
  var ImpressionsCellColor = "white";
  if(preImpressions >= currentSetting.minAlertImpressions || postImpressions >= currentSetting.minAlertImpressions) {
    if(currentSetting.minIncreaseForImpressionsAlert && ImpressionsChangePercent > 0 && ImpressionsChangePercent > currentSetting.minIncreaseForImpressionsAlert) {
      var alert = true;
      ImpressionsCellColor = currentSetting.colors["Impressions"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForImpressionsAlert && ImpressionsChangePercent < 0 && ImpressionsChangePercent < currentSetting.minDecreaseForImpressionsAlert) {
      var alert = true;
      ImpressionsCellColor = currentSetting.colors["Impressions"].bgColorForDecreases;
    } 
  }
  
  var preClicks = element.pre.Clicks;
    var postClicks = element.post.Clicks;
    if(preClicks > 0) {
      var ClicksChange = postClicks - preClicks;
      var ClicksChangePercent = (postClicks - preClicks) / preClicks;
      var ClicksTextColor = (ClicksChange > 0 ? currentSetting.colors["Clicks"].textColorForIncreases : currentSetting.colors["Clicks"].textColorForDecreases);
    } else if (preClicks == 0 && postClicks > 0) {
      var ClicksChange = postClicks;
      var ClicksChangePercent = "\u221E";
      var ClicksColor = currentSetting.colors["Clicks"].textColorForIncreases;
    } else if (preClicks == postClicks) {
      var ClicksChange = 0;
      var ClicksChangePercent = 0;
    }
  var ClicksCellColor = "white";
  if(preClicks >= currentSetting.minAlertClicks || postClicks >= currentSetting.minAlertClicks) {
    if(currentSetting.minIncreaseForClicksAlert && ClicksChangePercent > 0 && ClicksChangePercent > currentSetting.minIncreaseForClicksAlert) {
      var alert = true;
      ClicksCellColor = currentSetting.colors["Clicks"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForClicksAlert && ClicksChangePercent < 0 && ClicksChangePercent < currentSetting.minDecreaseForClicksAlert) {
      var alert = true;
      ClicksCellColor = currentSetting.colors["Clicks"].bgColorForDecreases;
    } 
  }
  
  var preAveragePosition = element.pre.AveragePosition;
    var postAveragePosition = element.post.AveragePosition;
    if(preAveragePosition > 0) {
      var AveragePositionChange = postAveragePosition - preAveragePosition;
      var AveragePositionChangePercent = (postAveragePosition - preAveragePosition) / preAveragePosition;
      var AveragePositionTextColor = (AveragePositionChange > 0 ? currentSetting.colors["AveragePosition"].textColorForIncreases : currentSetting.colors["AveragePosition"].textColorForDecreases);
    } else if (preAveragePosition == 0 && postAveragePosition > 0) {
      var AveragePositionChange = postAveragePosition;
      var AveragePositionChangePercent = "\u221E";
      var AveragePositionColor = currentSetting.colors["AveragePosition"].textColorForIncreases;
    } else if (preAveragePosition == postAveragePosition) {
      var AveragePositionChange = 0;
      var AveragePositionChangePercent = 0;
    }
  var AveragePositionCellColor = "white";
  if(preAveragePosition >= currentSetting.minAlertAveragePosition || postAveragePosition >= currentSetting.minAlertAveragePosition) {
    if(currentSetting.minIncreaseForAveragePositionAlert && AveragePositionChangePercent > 0 && AveragePositionChangePercent > currentSetting.minIncreaseForAveragePositionAlert) {
      var alert = true;
      AveragePositionCellColor = currentSetting.colors["AveragePosition"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForAveragePositionAlert && AveragePositionChangePercent < 0 && AveragePositionChangePercent < currentSetting.minDecreaseForAveragePositionAlert) {
      var alert = true;
      AveragePositionCellColor = currentSetting.colors["AveragePosition"].bgColorForDecreases;
    } 
  }
  
  var preCost = element.pre.Cost;
    var postCost = element.post.Cost;
    if(preCost > 0) {
      var CostChange = postCost - preCost;
      var CostChangePercent = (postCost - preCost) / preCost;
      var CostTextColor = (CostChange > 0 ? currentSetting.colors["Cost"].textColorForIncreases : currentSetting.colors["Cost"].textColorForDecreases);
    } else if (preCost == 0 && postCost > 0) {
      var CostChange = postCost;
      var CostChangePercent = "\u221E";
      var CostColor = currentSetting.colors["Cost"].textColorForIncreases;
    } else if (preCost == postCost) {
      var CostChange = 0;
      var CostChangePercent = 0;
    }
  var CostCellColor = "white";
  if(preCost >= currentSetting.minAlertCost || postCost >= currentSetting.minAlertCost) {
    if(currentSetting.minIncreaseForCostAlert && CostChangePercent > 0 && CostChangePercent > currentSetting.minIncreaseForCostAlert) {
      var alert = true;
      CostCellColor = currentSetting.colors["Cost"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForCostAlert && CostChangePercent < 0 && CostChangePercent < currentSetting.minDecreaseForCostAlert) {
      var alert = true;
      CostCellColor = currentSetting.colors["Cost"].bgColorForDecreases;
    } 
  }
  
  var preConversionRate = element.pre.ConversionRate;
    var postConversionRate = element.post.ConversionRate;
    if(preConversionRate > 0) {
      var ConversionRateChange = postConversionRate - preConversionRate;
      var ConversionRateChangePercent = (postConversionRate - preConversionRate) / preConversionRate;
      var ConversionRateTextColor = (ConversionRateChange > 0 ? currentSetting.colors["ConversionRate"].textColorForIncreases : currentSetting.colors["ConversionRate"].textColorForDecreases);
    } else if (preConversionRate == 0 && postConversionRate > 0) {
      var ConversionRateChange = postConversionRate;
      var ConversionRateChangePercent = "\u221E";
      var ConversionRateColor = currentSetting.colors["ConversionRate"].textColorForIncreases;
    } else if (preConversionRate == postConversionRate) {
      var ConversionRateChange = 0;
      var ConversionRateChangePercent = 0;
    }
  var ConversionRateCellColor = "white";
  if(preConversionRate >= currentSetting.minAlertConversionRate || postConversionRate >= currentSetting.minAlertConversionRate) {
    if(currentSetting.minIncreaseForConversionRateAlert && ConversionRateChangePercent > 0 && ConversionRateChangePercent > currentSetting.minIncreaseForConversionRateAlert) {
      var alert = true;
      ConversionRateCellColor = currentSetting.colors["ConversionRate"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForConversionRateAlert && ConversionRateChangePercent < 0 && ConversionRateChangePercent < currentSetting.minDecreaseForConversionRateAlert) {
      var alert = true;
      ConversionRateCellColor = currentSetting.colors["ConversionRate"].bgColorForDecreases;
    } 
  }
  
  var preConversions = element.pre.Conversions;
    var postConversions = element.post.Conversions;
  if(DEBUG) Logger.log("preConversions: " + preConversions);
  if(DEBUG) Logger.log("postConversions: " + postConversions);
    if(preConversions > 0) {
      var ConversionsChange = postConversions - preConversions;
      var ConversionsChangePercent = (postConversions - preConversions) / preConversions;
      var ConversionsTextColor = (ConversionsChange > 0 ? currentSetting.colors["Conversions"].textColorForIncreases : currentSetting.colors["Conversions"].textColorForDecreases);
    } else if (preConversions == 0 && postConversions > 0) {
      var ConversionsChange = postConversions;
      var ConversionsChangePercent = "\u221E";
      var ConversionsColor = currentSetting.colors["Conversions"].textColorForIncreases;
    } else if (preConversions == postConversions) {
      var ConversionsChange = 0;
      var ConversionsChangePercent = 0;
    }
  var ConversionsCellColor = "white";
  if(preConversions >= currentSetting.minAlertConversions || postConversions >= currentSetting.minAlertConversions) {
    if(currentSetting.minIncreaseForConversionsAlert && ConversionsChangePercent > 0 && ConversionsChangePercent > currentSetting.minIncreaseForConversionsAlert) {
      var alert = true;
      ConversionsCellColor = currentSetting.colors["Conversions"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForConversionsAlert && ConversionsChangePercent < 0 && ConversionsChangePercent < currentSetting.minDecreaseForConversionsAlert) {
      var alert = true;
      ConversionsCellColor = currentSetting.colors["Conversions"].bgColorForDecreases;
    } 
  }
  
  var preConversionValue = element.pre.ConversionValue;
    var postConversionValue = element.post.ConversionValue;
    if(preConversionValue > 0) {
      var ConversionValueChange = postConversionValue - preConversionValue;
      var ConversionValueChangePercent = (postConversionValue - preConversionValue) / preConversionValue;
      var ConversionValueTextColor = (ConversionValueChange > 0 ? currentSetting.colors["ConversionValue"].textColorForIncreases : currentSetting.colors["ConversionValue"].textColorForDecreases);
    } else if (preConversionValue == 0 && postConversionValue > 0) {
      var ConversionValueChange = postConversionValue;
      var ConversionValueChangePercent = "\u221E";
      var ConversionValueColor = currentSetting.colors["ConversionValue"].textColorForIncreases;
    } else if (preConversionValue == postConversionValue) {
      var ConversionValueChange = 0;
      var ConversionValueChangePercent = 0;
    }
  var ConversionValueCellColor = "white";
  if(preConversionValue >= currentSetting.minAlertConversionValue || postConversionValue >= currentSetting.minAlertConversionValue) {
    if(currentSetting.minIncreaseForConversionValueAlert && ConversionValueChangePercent > 0 && ConversionValueChangePercent > currentSetting.minIncreaseForConversionValueAlert) {
      var alert = true;
      ConversionValueCellColor = currentSetting.colors["ConversionValue"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForConversionValueAlert && ConversionValueChangePercent < 0 && ConversionValueChangePercent < currentSetting.minDecreaseForConversionValueAlert) {
      var alert = true;
      ConversionValueCellColor = currentSetting.colors["ConversionValue"].bgColorForDecreases;
    } 
  }
  /*
  var preConvertedClicks = element.pre.ConvertedClicks;
    var postConvertedClicks = element.post.ConvertedClicks;
    if(preConvertedClicks > 0) {
      var ConvertedClicksChange = postConvertedClicks - preConvertedClicks;
      var ConvertedClicksChangePercent = (postConvertedClicks - preConvertedClicks) / preConvertedClicks;
      var ConvertedClicksTextColor = (ConvertedClicksChange > 0 ? currentSetting.colors["ConvertedClicks"].textColorForIncreases : currentSetting.colors["ConvertedClicks"].textColorForDecreases);
    } else if (preConvertedClicks == 0 && postConvertedClicks > 0) {
      var ConvertedClicksChange = postConvertedClicks;
      var ConvertedClicksChangePercent = "\u221E";
      var ConvertedClicksColor = currentSetting.colors["ConvertedClicks"].textColorForIncreases;
    } else if (preConvertedClicks == postConvertedClicks) {
      var ConvertedClicksChange = 0;
      var ConvertedClicksChangePercent = 0;
    }
  
  if(preConvertedClicks >= currentSetting.minAlertConvertedClicks || postConvertedClicks >= currentSetting.minAlertConvertedClicks) {
    if(currentSetting.minIncreaseForConvertedClicksAlert && ConvertedClicksChangePercent > 0 && ConvertedClicksChangePercent > currentSetting.minIncreaseForConvertedClicksAlert) {
      var alert = true;
      ConvertedClicksCellColor = currentSetting.colors["ConvertedClicks"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForConvertedClicksAlert && ConvertedClicksChangePercent < 0 && ConvertedClicksChangePercent < currentSetting.minDecreaseForConvertedClicksAlert) {
      var alert = true;
      ConvertedClicksCellColor = currentSetting.colors["ConvertedClicks"].bgColorForDecreases;
    } 
  }*/
  
  var preCostPerConversion = element.pre.CostPerConversion;
    var postCostPerConversion = element.post.CostPerConversion;
    if(preCostPerConversion > 0) {
      var CostPerConversionChange = postCostPerConversion - preCostPerConversion;
      var CostPerConversionChangePercent = (postCostPerConversion - preCostPerConversion) / preCostPerConversion;
      var CostPerConversionTextColor = (CostPerConversionChange > 0 ? currentSetting.colors["CostPerConversion"].textColorForIncreases : currentSetting.colors["CostPerConversion"].textColorForDecreases);
    } else if (preCostPerConversion == 0 && postCostPerConversion > 0) {
      var CostPerConversionChange = postCostPerConversion;
      var CostPerConversionChangePercent = "\u221E";
      var CostPerConversionColor = currentSetting.colors["CostPerConversion"].textColorForIncreases;
    } else if (preCostPerConversion == postCostPerConversion) {
      var CostPerConversionChange = 0;
      var CostPerConversionChangePercent = 0;
    }
  var CostPerConversionCellColor = "white";
  if(preCostPerConversion >= currentSetting.minAlertCostPerConversion || postCostPerConversion >= currentSetting.minAlertCostPerConversion) {
    if(currentSetting.minIncreaseForCostPerConversionAlert && CostPerConversionChangePercent > 0 && CostPerConversionChangePercent > currentSetting.minIncreaseForCostPerConversionAlert) {
      var alert = true;
      CostPerConversionCellColor = currentSetting.colors["CostPerConversion"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForCostPerConversionAlert && CostPerConversionChangePercent < 0 && CostPerConversionChangePercent < currentSetting.minDecreaseForCostPerConversionAlert) {
      var alert = true;
      CostPerConversionCellColor = currentSetting.colors["CostPerConversion"].bgColorForDecreases;
    } 
  }
  
  var preCostPerAllConversion = element.pre.CostPerAllConversion;
    var postCostPerAllConversion = element.post.CostPerAllConversion;
    if(preCostPerAllConversion > 0) {
      var CostPerAllConversionChange = postCostPerAllConversion - preCostPerAllConversion;
      var CostPerAllConversionChangePercent = (postCostPerAllConversion - preCostPerAllConversion) / preCostPerAllConversion;
      var CostPerAllConversionTextColor = (CostPerAllConversionChange > 0 ? currentSetting.colors["CostPerAllConversion"].textColorForIncreases : currentSetting.colors["CostPerAllConversion"].textColorForDecreases);
    } else if (preCostPerAllConversion == 0 && postCostPerAllConversion > 0) {
      var CostPerAllConversionChange = postCostPerAllConversion;
      var CostPerAllConversionChangePercent = "\u221E";
      var CostPerAllConversionColor = currentSetting.colors["CostPerAllConversion"].textColorForIncreases;
    } else if (preCostPerAllConversion == postCostPerAllConversion) {
      var CostPerAllConversionChange = 0;
      var CostPerAllConversionChangePercent = 0;
    }
  var CostPerAllConversionCellColor = "white";
  if(preCostPerAllConversion >= currentSetting.minAlertCostPerAllConversion || postCostPerAllConversion >= currentSetting.minAlertCostPerAllConversion) {
    if(currentSetting.minIncreaseForCostPerAllConversionAlert && CostPerAllConversionChangePercent > 0 && CostPerAllConversionChangePercent > currentSetting.minIncreaseForCostPerAllConversionAlert) {
      var alert = true;
      CostPerAllConversionCellColor = currentSetting.colors["CostPerAllConversion"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForCostPerAllConversionAlert && CostPerAllConversionChangePercent < 0 && CostPerAllConversionChangePercent < currentSetting.minDecreaseForCostPerAllConversionAlert) {
      var alert = true;
      CostPerAllConversionCellColor = currentSetting.colors["CostPerAllConversion"].bgColorForDecreases;
    } 
  }
  
  var preCrossDeviceConversions = element.pre.CrossDeviceConversions;
    var postCrossDeviceConversions = element.post.CrossDeviceConversions;
    if(preCrossDeviceConversions > 0) {
      var CrossDeviceConversionsChange = postCrossDeviceConversions - preCrossDeviceConversions;
      var CrossDeviceConversionsChangePercent = (postCrossDeviceConversions - preCrossDeviceConversions) / preCrossDeviceConversions;
      var CrossDeviceConversionsTextColor = (CrossDeviceConversionsChange > 0 ? currentSetting.colors["CrossDeviceConversions"].textColorForIncreases : currentSetting.colors["CrossDeviceConversions"].textColorForDecreases);
    } else if (preCrossDeviceConversions == 0 && postCrossDeviceConversions > 0) {
      var CrossDeviceConversionsChange = postCrossDeviceConversions;
      var CrossDeviceConversionsChangePercent = "\u221E";
      var CrossDeviceConversionsColor = currentSetting.colors["CrossDeviceConversions"].textColorForIncreases;
    } else if (preCrossDeviceConversions == postCrossDeviceConversions) {
      var CrossDeviceConversionsChange = 0;
      var CrossDeviceConversionsChangePercent = 0;
    }
  var CrossDeviceConversionsCellColor = "white";
  if(preCrossDeviceConversions >= currentSetting.minAlertCrossDeviceConversions || postCrossDeviceConversions >= currentSetting.minAlertCrossDeviceConversions) {
    if(currentSetting.minIncreaseForCrossDeviceConversionsAlert && CrossDeviceConversionsChangePercent > 0 && CrossDeviceConversionsChangePercent > currentSetting.minIncreaseForCrossDeviceConversionsAlert) {
      var alert = true;
      CrossDeviceConversionsCellColor = currentSetting.colors["CrossDeviceConversions"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForCrossDeviceConversionsAlert && CrossDeviceConversionsChangePercent < 0 && CrossDeviceConversionsChangePercent < currentSetting.minDecreaseForCrossDeviceConversionsAlert) {
      var alert = true;
      CrossDeviceConversionsCellColor = currentSetting.colors["CrossDeviceConversions"].bgColorForDecreases;
    } 
  }
  
  var preValuePerConversion = element.pre.ValuePerConversion;
    var postValuePerConversion = element.post.ValuePerConversion;
    if(preValuePerConversion > 0) {
      var ValuePerConversionChange = postValuePerConversion - preValuePerConversion;
      var ValuePerConversionChangePercent = (postValuePerConversion - preValuePerConversion) / preValuePerConversion;
      var ValuePerConversionTextColor = (ValuePerConversionChange > 0 ? currentSetting.colors["ValuePerConversion"].textColorForIncreases : currentSetting.colors["ValuePerConversion"].textColorForDecreases);
    } else if (preValuePerConversion == 0 && postValuePerConversion > 0) {
      var ValuePerConversionChange = postValuePerConversion;
      var ValuePerConversionChangePercent = "\u221E";
      var ValuePerConversionColor = currentSetting.colors["ValuePerConversion"].textColorForIncreases;
    } else if (preValuePerConversion == postValuePerConversion) {
      var ValuePerConversionChange = "-";
      var ValuePerConversionChangePercent = 0;
    }
  var ValuePerConversionCellColor = "white";
  if(preValuePerConversion >= currentSetting.minAlertValuePerConversion || postValuePerConversion >= currentSetting.minAlertValuePerConversion) {
    if(currentSetting.minIncreaseForValuePerConversionAlert && ValuePerConversionChangePercent > 0 && ValuePerConversionChangePercent > currentSetting.minIncreaseForValuePerConversionAlert) {
      var alert = true;
      ValuePerConversionCellColor = currentSetting.colors["ValuePerConversion"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForValuePerConversionAlert && ValuePerConversionChangePercent < 0 && ValuePerConversionChangePercent < currentSetting.minDecreaseForValuePerConversionAlert) {
      var alert = true;
      ValuePerConversionCellColor = currentSetting.colors["ValuePerConversion"].bgColorForDecreases;
    } 
  }
  
  var preValuePerAllConversion = element.pre.ValuePerAllConversion;
    var postValuePerAllConversion = element.post.ValuePerAllConversion;
    if(preValuePerAllConversion > 0) {
      var ValuePerAllConversionChange = postValuePerAllConversion - preValuePerAllConversion;
      var ValuePerAllConversionChangePercent = (postValuePerAllConversion - preValuePerAllConversion) / preValuePerAllConversion;
      var ValuePerAllConversionTextColor = (ValuePerAllConversionChange > 0 ? currentSetting.colors["ValuePerAllConversion"].textColorForIncreases : currentSetting.colors["ValuePerAllConversion"].textColorForDecreases);
    } else if (preValuePerAllConversion == 0 && postValuePerAllConversion > 0) {
      var ValuePerAllConversionChange = postValuePerAllConversion;
      var ValuePerAllConversionChangePercent = "\u221E";
      var ValuePerAllConversionColor = currentSetting.colors["ValuePerAllConversion"].textColorForIncreases;
    } else if (preValuePerAllConversion == postValuePerAllConversion) {
      var ValuePerAllConversionChange = 0;
      var ValuePerAllConversionChangePercent = 0;
    }
  var ValuePerAllConversionCellColor = "white";
  if(preValuePerAllConversion >= currentSetting.minAlertValuePerAllConversion || postValuePerAllConversion >= currentSetting.minAlertValuePerAllConversion) {
    if(currentSetting.minIncreaseForValuePerAllConversionAlert && ValuePerAllConversionChangePercent > 0 && ValuePerAllConversionChangePercent > currentSetting.minIncreaseForValuePerAllConversionAlert) {
      var alert = true;
      ValuePerAllConversionCellColor = currentSetting.colors["ValuePerAllConversion"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForValuePerAllConversionAlert && ValuePerAllConversionChangePercent < 0 && ValuePerAllConversionChangePercent < currentSetting.minDecreaseForValuePerAllConversionAlert) {
      var alert = true;
      ValuePerAllConversionCellColor = currentSetting.colors["ValuePerAllConversion"].bgColorForDecreases;
    } 
  }
  
  var preViewThroughConversions = element.pre.ViewThroughConversions;
    var postViewThroughConversions = element.post.ViewThroughConversions;
    if(preViewThroughConversions > 0) {
      var ViewThroughConversionsChange = postViewThroughConversions - preViewThroughConversions;
      var ViewThroughConversionsChangePercent = (postViewThroughConversions - preViewThroughConversions) / preViewThroughConversions;
      var ViewThroughConversionsTextColor = (ViewThroughConversionsChange > 0 ? currentSetting.colors["ViewThroughConversions"].textColorForIncreases : currentSetting.colors["ViewThroughConversions"].textColorForDecreases);
    } else if (preViewThroughConversions == 0 && postViewThroughConversions > 0) {
      var ViewThroughConversionsChange = postViewThroughConversions;
      var ViewThroughConversionsChangePercent = "\u221E";
      var ViewThroughConversionsColor = currentSetting.colors["ViewThroughConversions"].textColorForIncreases;
    } else if (preViewThroughConversions == postViewThroughConversions) {
      var ViewThroughConversionsChange = 0;
      var ViewThroughConversionsChangePercent = 0;
    }
  var ViewThroughConversionsCellColor = "white";
  if(preViewThroughConversions >= currentSetting.minAlertViewThroughConversions || postViewThroughConversions >= currentSetting.minAlertViewThroughConversions) {
    if(currentSetting.minIncreaseForViewThroughConversionsAlert && ViewThroughConversionsChangePercent > 0 && ViewThroughConversionsChangePercent > currentSetting.minIncreaseForViewThroughConversionsAlert) {
      var alert = true;
      ViewThroughConversionsCellColor = currentSetting.colors["ViewThroughConversions"].bgColorForIncreases;
    }
    if (currentSetting.minDecreaseForViewThroughConversionsAlert && ViewThroughConversionsChangePercent < 0 && ViewThroughConversionsChangePercent < currentSetting.minDecreaseForViewThroughConversionsAlert) {
      var alert = true;
      ViewThroughConversionsCellColor = currentSetting.colors["ViewThroughConversions"].bgColorForDecreases;
    } 
  }
  
  
    
  if(alert) {
    var accountName = element.accountName;
    var campaignName = element.campaignName;
    var adGroupName = element.adGroupName;
    var keyword = element.keyword;
    var headline = element.headline;
    var line1 = element.line1;
    var line2 = element.line2;
    var visibleUrl = element.visibleUrl;
    var adId = element.adId;
    
    
    sheet.getRange("A" + lineCounter).setValue(accountName);
    sheet.getRange("B" + lineCounter).setValue(campaignName);
    sheet.getRange("C" + lineCounter).setValue(adGroupName);
    sheet.getRange("D" + lineCounter).setValue(keyword);
    
    if(scope == "Ad") {
      sheet.getRange("E" + lineCounter).setValue(headline);
      sheet.getRange("F" + lineCounter).setValue(line1);
      sheet.getRange("G" + lineCounter).setValue(line2);
      sheet.getRange("H" + lineCounter).setValue(visibleUrl);
      sheet.getRange("I" + lineCounter).setValue(adId);
      sheet.getRange("J" + lineCounter).setValue(segment1);
      sheet.getRange("K" + lineCounter).setValue(segment2);
      sheet.getRange("L" + lineCounter).setValue(segment3);
      sheet.getRange("M" + lineCounter).setValue(segment4);
      
      sheet.getRange("N" + lineCounter).setValue(preAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("O" + lineCounter).setValue(postAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("P" + lineCounter).setValue(AllConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionsColor);
      
      sheet.getRange("Q" + lineCounter).setValue(preAllConversionValue).setBackground(AllConversionValueCellColor);
      sheet.getRange("R" + lineCounter).setValue(postAllConversionValue).setBackground(AllConversionValueCellColor);
      sheet.getRange("S" + lineCounter).setValue(AllConversionValueChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionValueColor);
      
      sheet.getRange("T" + lineCounter).setValue(preAllConversionRate).setBackground(AllConversionRateCellColor);
      sheet.getRange("U" + lineCounter).setValue(postAllConversionRate).setBackground(AllConversionRateCellColor);
      sheet.getRange("V" + lineCounter).setValue(AllConversionRateChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionRateTextColor);
      
      sheet.getRange("W" + lineCounter).setValue(preAverageCpc).setBackground(AverageCpcCellColor);
      sheet.getRange("X" + lineCounter).setValue(postAverageCpc).setBackground(AverageCpcCellColor);
      sheet.getRange("Y" + lineCounter).setValue(AverageCpcChangePercent).setNumberFormat("#0.00%").setFontColor(AverageCpcTextColor);
      
    /*  sheet.getRange("Z" + lineCounter).setValue(preCtr).setBackground(CtrCellColor);
      sheet.getRange("AA" + lineCounter).setValue(postCtr).setBackground(CtrCellColor);
      sheet.getRange("AB" + lineCounter).setValue(CtrChangePercent).setNumberFormat("#0.00%").setFontColor(CtrTextColor);
      
      sheet.getRange("AC" + lineCounter).setValue(preImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("AD" + lineCounter).setValue(postImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("AE" + lineCounter).setValue(ImpressionsChangePercent).setNumberFormat("#0.00%").setFontColor(ImpressionsTextColor);
      
      sheet.getRange("AF" + lineCounter).setValue(preClicks).setBackground(ClicksCellColor);
      sheet.getRange("AG" + lineCounter).setValue(postClicks).setBackground(ClicksCellColor);
      sheet.getRange("AH" + lineCounter).setValue(ClicksChangePercent).setNumberFormat("#0.00%").setFontColor(ClicksTextColor);
      
      sheet.getRange("AI" + lineCounter).setValue(preAveragePosition).setBackground(AveragePositionCellColor);
      sheet.getRange("AJ" + lineCounter).setValue(postAveragePosition).setBackground(AveragePositionCellColor);
      sheet.getRange("AK" + lineCounter).setValue(AveragePositionChangePercent).setNumberFormat("#0.00%").setFontColor(AveragePositionTextColor);
      
      sheet.getRange("AL" + lineCounter).setValue(preCost).setBackground(CostCellColor);
      sheet.getRange("AM" + lineCounter).setValue(postCost).setBackground(CostCellColor);
      sheet.getRange("AN" + lineCounter).setValue(CostChangePercent).setNumberFormat("#0.00%").setFontColor(CostTextColor);
      
      sheet.getRange("AO" + lineCounter).setValue(preConversionRate).setBackground(ConversionRateCellColor);
      sheet.getRange("AP" + lineCounter).setValue(postConversionRate).setBackground(ConversionRateCellColor);
      sheet.getRange("AQ" + lineCounter).setValue(ConversionRateChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionRateTextColor);
      
      sheet.getRange("AR" + lineCounter).setValue(preConversions).setBackground(ConversionsCellColor);
      sheet.getRange("AS" + lineCounter).setValue(postConversions).setBackground(ConversionsCellColor);
      sheet.getRange("AT" + lineCounter).setValue(ConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionsTextColor);
      
      sheet.getRange("AU" + lineCounter).setValue(preConversionValue).setBackground(ConversionValueCellColor);
      sheet.getRange("AV" + lineCounter).setValue(postConversionValue).setBackground(ConversionValueCellColor);
      sheet.getRange("AW" + lineCounter).setValue(ConversionValueChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionValueTextColor);
      
      sheet.getRange("AX" + lineCounter).setValue("NANA");
      sheet.getRange("AY" + lineCounter).setValue("NANA");
      sheet.getRange("AZ" + lineCounter).setValue("NANA");
      
      sheet.getRange("BA" + lineCounter).setValue(preCostPerConversion).setBackground(CostPerConversionCellColor);
      sheet.getRange("BB" + lineCounter).setValue(postCostPerConversion).setBackground(CostPerConversionCellColor);
      sheet.getRange("BC" + lineCounter).setValue(CostPerConversionChangePercent).setNumberFormat("#0.00%").setFontColor(CostPerConversionTextColor);
      
      sheet.getRange("BD" + lineCounter).setValue(preCostPerAllConversion).setBackground(CostPerAllConversionCellColor);
      sheet.getRange("BE" + lineCounter).setValue(postCostPerAllConversion).setBackground(CostPerAllConversionCellColor);
      sheet.getRange("BF" + lineCounter).setValue(CostPerAllConversionChangePercent).setNumberFormat("#0.00%").setFontColor(CostPerAllConversionTextColor);
      
      sheet.getRange("BG" + lineCounter).setValue(preCrossDeviceConversions).setBackground(CrossDeviceConversionsCellColor);
      sheet.getRange("BH" + lineCounter).setValue(postCrossDeviceConversions).setBackground(CrossDeviceConversionsCellColor);
      sheet.getRange("BI" + lineCounter).setValue(CrossDeviceConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(CrossDeviceConversionsTextColor);
      
      sheet.getRange("BJ" + lineCounter).setValue(preValuePerConversion).setBackground(ValuePerConversionCellColor);
      sheet.getRange("BK" + lineCounter).setValue(postValuePerConversion).setBackground(ValuePerConversionCellColor);
      sheet.getRange("BL" + lineCounter).setValue(ValuePerConversionChangePercent).setNumberFormat("#0.00%").setFontColor(ValuePerConversionTextColor);
      
      sheet.getRange("BM" + lineCounter).setValue(preValuePerAllConversion).setBackground(ValuePerAllConversionCellColor);
      sheet.getRange("BN" + lineCounter).setValue(postValuePerAllConversion).setBackground(ValuePerAllConversionCellColor);
      sheet.getRange("BO" + lineCounter).setValue(ValuePerAllConversionChangePercent).setNumberFormat("#0.00%").setFontColor(ValuePerAllConversionTextColor);
      
      sheet.getRange("BP" + lineCounter).setValue(preViewThroughConversions).setBackground(ViewThroughConversionsCellColor);
      sheet.getRange("BQ" + lineCounter).setValue(postViewThroughConversions).setBackground(ViewThroughConversionsCellColor);
      sheet.getRange("BR" + lineCounter).setValue(ViewThroughConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(ViewThroughConversionsTextColor);
*/

      
     /* sheet.getRange("BS" + lineCounter).setValue(scope);*/
    } else {
    


    /*our data dump*/
    sheet.getRange("E" + lineCounter).setValue(segment1);
      sheet.getRange("F" + lineCounter).setValue(segment2);
      sheet.getRange("G" + lineCounter).setValue(segment3);
      sheet.getRange("H" + lineCounter).setValue(segment4);

/*all conversions*/
   /*   sheet.getRange("I" + lineCounter).setValue(preAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("J" + lineCounter).setValue(postAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("K" + lineCounter).setValue(AllConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionsColor);*/
      
/*Impressions was x y z*/
      sheet.getRange("R" + lineCounter).setValue(preImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("S" + lineCounter).setValue(postImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("T" + lineCounter).setValue(ImpressionsChangePercent).setNumberFormat("#0.00%").setFontColor(ImpressionsTextColor);

/*clicks changed from aa ab ac to L  J K*/
      sheet.getRange("L" + lineCounter).setValue(preClicks).setBackground(ClicksCellColor);
      sheet.getRange("M" + lineCounter).setValue(postClicks).setBackground(ClicksCellColor);
      sheet.getRange("N" + lineCounter).setValue(ClicksChangePercent).setNumberFormat("#0.00%").setFontColor(ClicksTextColor);
     
/*cost was ag ah ai */
     sheet.getRange("O" + lineCounter).setValue(preCost).setBackground(CostCellColor);
    sheet.getRange("P" + lineCounter).setValue(postCost).setBackground(CostCellColor);
     sheet.getRange("Q" + lineCounter).setValue(CostChangePercent).setNumberFormat("#0.00%").setFontColor(CostTextColor);

     /*conversions*/
      sheet.getRange("I" + lineCounter).setValue(preConversions).setBackground(ConversionsCellColor);
      sheet.getRange("J" + lineCounter).setValue(postConversions).setBackground(ConversionsCellColor);
      sheet.getRange("K" + lineCounter).setValue(ConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionsTextColor);


     /* sheet.getRange("E" + lineCounter).setValue(segment1);
      sheet.getRange("F" + lineCounter).setValue(segment2);
      sheet.getRange("G" + lineCounter).setValue(segment3);
      sheet.getRange("H" + lineCounter).setValue(segment4);
      
      sheet.getRange("I" + lineCounter).setValue(preAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("J" + lineCounter).setValue(postAllConversions).setBackground(AllConversionValueCellColor);
      sheet.getRange("K" + lineCounter).setValue(AllConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionsColor);
      
      sheet.getRange("L" + lineCounter).setValue(preAllConversionValue).setBackground(AllConversionValueCellColor);
      sheet.getRange("M" + lineCounter).setValue(postAllConversionValue).setBackground(AllConversionValueCellColor);
      sheet.getRange("N" + lineCounter).setValue(AllConversionValueChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionValueColor);
      
      sheet.getRange("O" + lineCounter).setValue(preAllConversionRate).setBackground(AllConversionRateCellColor);
      sheet.getRange("P" + lineCounter).setValue(postAllConversionRate).setBackground(AllConversionRateCellColor);
      sheet.getRange("Q" + lineCounter).setValue(AllConversionRateChangePercent).setNumberFormat("#0.00%").setFontColor(AllConversionRateTextColor);
      
      sheet.getRange("R" + lineCounter).setValue(preAverageCpc).setBackground(AverageCpcCellColor);
      sheet.getRange("S" + lineCounter).setValue(postAverageCpc).setBackground(AverageCpcCellColor);
      sheet.getRange("T" + lineCounter).setValue(AverageCpcChangePercent).setNumberFormat("#0.00%").setFontColor(AverageCpcTextColor);
    
      sheet.getRange("U" + lineCounter).setValue(preCtr).setBackground(CtrCellColor);
      sheet.getRange("V" + lineCounter).setValue(postCtr).setBackground(CtrCellColor);
      sheet.getRange("W" + lineCounter).setValue(CtrChangePercent).setNumberFormat("#0.00%").setFontColor(CtrTextColor);
      
      sheet.getRange("X" + lineCounter).setValue(preImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("Y" + lineCounter).setValue(postImpressions).setBackground(ImpressionsCellColor);
      sheet.getRange("Z" + lineCounter).setValue(ImpressionsChangePercent).setNumberFormat("#0.00%").setFontColor(ImpressionsTextColor);
      
      sheet.getRange("AA" + lineCounter).setValue(preClicks).setBackground(ClicksCellColor);
      sheet.getRange("AB" + lineCounter).setValue(postClicks).setBackground(ClicksCellColor);
      sheet.getRange("AC" + lineCounter).setValue(ClicksChangePercent).setNumberFormat("#0.00%").setFontColor(ClicksTextColor);
     
      sheet.getRange("AD" + lineCounter).setValue(preAveragePosition).setBackground(AveragePositionCellColor);
      sheet.getRange("AE" + lineCounter).setValue(postAveragePosition).setBackground(AveragePositionCellColor);
      sheet.getRange("AF" + lineCounter).setValue(AveragePositionChangePercent).setNumberFormat("#0.00%").setFontColor(AveragePositionTextColor);
      
      sheet.getRange("AG" + lineCounter).setValue(preCost).setBackground(CostCellColor);
      sheet.getRange("AH" + lineCounter).setValue(postCost).setBackground(CostCellColor);
      sheet.getRange("AI" + lineCounter).setValue(CostChangePercent).setNumberFormat("#0.00%").setFontColor(CostTextColor);
      
      sheet.getRange("AJ" + lineCounter).setValue(preConversionRate).setBackground(ConversionRateCellColor);
      sheet.getRange("AK" + lineCounter).setValue(postConversionRate).setBackground(ConversionRateCellColor);
      sheet.getRange("AL" + lineCounter).setValue(ConversionRateChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionRateTextColor);
      
      sheet.getRange("AM" + lineCounter).setValue(preConversions).setBackground(ConversionsCellColor);
      sheet.getRange("AN" + lineCounter).setValue(postConversions).setBackground(ConversionsCellColor);
      sheet.getRange("AO" + lineCounter).setValue(ConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionsTextColor);
      
      sheet.getRange("AP" + lineCounter).setValue(preConversionValue).setBackground(ConversionValueCellColor);
      sheet.getRange("AQ" + lineCounter).setValue(postConversionValue).setBackground(ConversionValueCellColor);
      sheet.getRange("AR" + lineCounter).setValue(ConversionValueChangePercent).setNumberFormat("#0.00%").setFontColor(ConversionValueTextColor);
      
      sheet.getRange("AS" + lineCounter).setValue("NANA");
      sheet.getRange("AT" + lineCounter).setValue("NANA");
      sheet.getRange("AU" + lineCounter).setValue("NANA");
      
      sheet.getRange("AV" + lineCounter).setValue(preCostPerConversion).setBackground(CostPerConversionCellColor);
      sheet.getRange("AW" + lineCounter).setValue(postCostPerConversion).setBackground(CostPerConversionCellColor);
      sheet.getRange("AX" + lineCounter).setValue(CostPerConversionChangePercent).setNumberFormat("#0.00%").setFontColor(CostPerConversionTextColor);
      
      sheet.getRange("AY" + lineCounter).setValue(preCostPerAllConversion).setBackground(CostPerAllConversionCellColor);
      sheet.getRange("AZ" + lineCounter).setValue(postCostPerAllConversion).setBackground(CostPerAllConversionCellColor);
      sheet.getRange("BA" + lineCounter).setValue(CostPerAllConversionChangePercent).setNumberFormat("#0.00%").setFontColor(CostPerAllConversionTextColor);
      
      sheet.getRange("BB" + lineCounter).setValue(preCrossDeviceConversions).setBackground(CrossDeviceConversionsCellColor);
      sheet.getRange("BC" + lineCounter).setValue(postCrossDeviceConversions).setBackground(CrossDeviceConversionsCellColor);
      sheet.getRange("BD" + lineCounter).setValue(CrossDeviceConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(CrossDeviceConversionsTextColor);
     
      sheet.getRange("BE" + lineCounter).setValue(preValuePerConversion).setBackground(ValuePerConversionCellColor);
      sheet.getRange("BF" + lineCounter).setValue(postValuePerConversion).setBackground(ValuePerConversionCellColor);
      sheet.getRange("BG" + lineCounter).setValue(ValuePerConversionChangePercent).setNumberFormat("#0.00%").setFontColor(ValuePerConversionTextColor);
      
      sheet.getRange("BH" + lineCounter).setValue(preValuePerAllConversion).setBackground(ValuePerAllConversionCellColor);
      sheet.getRange("BI" + lineCounter).setValue(postValuePerAllConversion).setBackground(ValuePerAllConversionCellColor);
      sheet.getRange("BJ" + lineCounter).setValue(ValuePerAllConversionChangePercent).setNumberFormat("#0.00%").setFontColor(ValuePerAllConversionTextColor);
      
      sheet.getRange("BK" + lineCounter).setValue(preViewThroughConversions).setBackground(ViewThroughConversionsCellColor);
      sheet.getRange("BL" + lineCounter).setValue(postViewThroughConversions).setBackground(ViewThroughConversionsCellColor);
      sheet.getRange("BM" + lineCounter).setValue(ViewThroughConversionsChangePercent).setNumberFormat("#0.00%").setFontColor(ViewThroughConversionsTextColor);
      
      sheet.getRange("BN" + lineCounter).setValue(scope);*/
      
      /*
      sheet.getRange("I" + lineCounter).setValue(preConversions);
      sheet.getRange("J" + lineCounter).setValue(postConversions);
      sheet.getRange("K" + lineCounter).setValue(conversionChange).setNumberFormat("#0.00%").setFontColor(conversionColor);
      sheet.getRange("L" + lineCounter).setValue(preClicks);
      sheet.getRange("M" + lineCounter).setValue(postClicks);
      sheet.getRange("N" + lineCounter).setValue(clickChange).setNumberFormat("#0.00%").setFontColor(clickColor);
      sheet.getRange("O" + lineCounter).setValue(preImpressions).setBackground(impressionsCellColor);;
      sheet.getRange("P" + lineCounter).setValue(postImpressions).setBackground(impressionsCellColor);;
      sheet.getRange("Q" + lineCounter).setValue(impressionsChangePercent).setNumberFormat("#0.00%").setFontColor(impressionsTextColor);
      sheet.getRange("R" + lineCounter).setValue(preCost);
      sheet.getRange("S" + lineCounter).setValue(postCost);
      sheet.getRange("T" + lineCounter).setValue(costChange).setNumberFormat("#0.00%").setFontColor(costColor);
      sheet.getRange("U" + lineCounter).setValue(preAvgPos);
      sheet.getRange("V" + lineCounter).setValue(postAvgPos);
      sheet.getRange("W" + lineCounter).setValue(avgPosChange).setNumberFormat("#0.00%").setFontColor(avgPosColor);
      sheet.getRange("X" + lineCounter).setValue(preConversionRate);
      sheet.getRange("Y" + lineCounter).setValue(postConversionRate);
      sheet.getRange("Z" + lineCounter).setValue(conversionRateChange).setNumberFormat("#0.00%").setFontColor(conversionRateColor);
      sheet.getRange("AA" + lineCounter).setValue(preCtr);
      sheet.getRange("AB" + lineCounter).setValue(postCtr);
      sheet.getRange("AC" + lineCounter).setValue(ctrChange).setNumberFormat("#0.00%").setFontColor(ctrColor);
      sheet.getRange("AD" + lineCounter).setValue(preAvgCpc);
      sheet.getRange("AE" + lineCounter).setValue(postAvgCpc);
      sheet.getRange("AF" + lineCounter).setValue(avgCpcChange).setNumberFormat("#0.00%").setFontColor(avgCpcColor);
      sheet.getRange("AG" + lineCounter).setValue(preCPA).setNumberFormat("0.00");
      sheet.getRange("AH" + lineCounter).setValue(postCPA).setNumberFormat("0.00");
      sheet.getRange("AI" + lineCounter).setValue(cpaChange).setNumberFormat("#0.00%").setFontColor(cpaColor);
      sheet.getRange("AJ" + lineCounter).setValue(preConversionValue).setNumberFormat("0.00");
      sheet.getRange("AK" + lineCounter).setValue(postConversionValue).setNumberFormat("0.00");
      sheet.getRange("AL" + lineCounter).setValue(conversionValueChange).setNumberFormat("#0.00%").setFontColor(conversionValueColor);
      sheet.getRange("AM" + lineCounter).setValue(preConversionValuePerClick).setNumberFormat("0.00");
      sheet.getRange("AN" + lineCounter).setValue(postConversionValuePerClick).setNumberFormat("0.00");
      sheet.getRange("AO" + lineCounter).setValue(conversionValuePerClickChange).setNumberFormat("#0.00%").setFontColor(conversionValuePerClickColor);
      sheet.getRange("AP" + lineCounter).setValue(scope);
      */
    } 
  }
  return(alert);
}





Date.prototype.yyyymmdd = function() {
    var yyyy = this.getFullYear().toString();
    var mm = (this.getMonth()+1).toString(); // getMonth() is zero-based
    var dd  = this.getDate().toString();
    return yyyy + (mm[1]?mm:"0"+mm[0]) + (dd[1]?dd:"0"+dd[0]); // padding
  };





  
  function sendEmailNotifications(emailAddresses, subject, body, emailType ) {
  
    if(emailType.toLowerCase().indexOf("warning") != -1) {
      var finalSubject = "[Warning] " + subject + " - " + AdWordsApp.currentAccount().getName() + " (" + AdWordsApp.currentAccount().getCustomerId() + ")"
    } else if(emailType.toLowerCase().indexOf("notification") != -1) {
      var finalSubject = "[Notification] " + subject + " - " + AdWordsApp.currentAccount().getName() + " (" + AdWordsApp.currentAccount().getCustomerId() + ")"
    }
    
    if(AdWordsApp.getExecutionInfo().isPreview()) {
      var finalBody = "<b>This script ran in preview mode. No changes were made to your account.</b><br/>" + body;
    } else {
      var finalBody = body;
    }
    
  MailApp.sendEmail({
        to:emailAddresses, 
        subject:  finalSubject,
        htmlBody: finalBody
      });
    
    if(DEBUG == 1) Logger.log("email sent to " + emailAddresses + ": " + finalSubject);

  }
  
 
  Date.prototype.yyyymmdd = function() {
  var yyyy = this.getFullYear().toString();
  var mm = (this.getMonth()+1).toString(); // getMonth() is zero-based
  var dd  = this.getDate().toString();
  return yyyy + (mm[1]?mm:"0"+mm[0]) + (dd[1]?dd:"0"+dd[0]); // padding
};
  
  
  
  function getFloat (input) {
    if(!input || input == "" || typeof(input) === 'undefined') var input = "0.0";
    input = input.toString();
    var output = parseFloat(input.replace(/,/g, ""));
    return output;
  }
  

  function setUpReportInGoogleSheets(spreadsheetUrl, spreadsheetName, accountManagers, overWriteOldData, sheetNames, folderNames) {
    
    var destinationSpreadsheet = new Object();
    if(folderNames) {
      var folderStructure = folderNames.split(",");
    } else {
      var folderStructure = new Array();
    }
    
    
    var targetFolder = DriveApp.getRootFolder();
    for(var i = 0; i < folderStructure.length; i++) {
      var folderName = folderStructure[i];
      if(folderName.toLowerCase().indexOf("[account id]") != -1) {
        folderName = AdWordsApp.currentAccount().getCustomerId();
      } else if(folderName.toLowerCase().indexOf("[account name]") != -1) {
        folderName = AdWordsApp.currentAccount().getName();
      }
      Logger.log("folderName: " + folderName);
      var foldersIterator = targetFolder.getFoldersByName(folderName);
      if (foldersIterator.hasNext()) {
        
        targetFolder = foldersIterator.next();
        Logger.log("Selected target folder: " + folderName);
      } else {
        if(DEBUG==1) Logger.log("Creating a new folder: " + folderName);
        targetFolder = targetFolder.createFolder(folderName);
      }
    }
    
    
    destinationSpreadsheet.overWrite = overWriteOldData;
    
    if(!spreadsheetUrl || spreadsheetUrl == "" || spreadsheetUrl == " " || spreadsheetUrl.toLowerCase().indexOf("new") != -1) var isNew = 1;
    destinationSpreadsheet.isNew = isNew;
    
    if(!sheetNames || !sheetNames[0]) {
      var sheetNames = new Array();
      sheetNames[0] = "Sheet 1";
    }
    
    if(isNew)
    {
      var spreadsheet = SpreadsheetApp.create(spreadsheetName);
      var id = spreadsheet.getId();
      var spreadsheetUrl = spreadsheet.getUrl();
      var file = DriveApp.getFileById(id);
      targetFolder.addFile(file);
      if(folderName) DriveApp.getRootFolder().removeFile(file);
    } 
    var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    destinationSpreadsheet.spreadsheet = spreadsheet;
    destinationSpreadsheet.url = spreadsheet.getUrl();
    
    // IF NEW -> REMOVE ALL SHEETS, THEN CREATE ALL SHEETS
    if(isNew){
      var allSheets = spreadsheet.getSheets(); 
      
      // remove
      for(var i=1,len=allSheets.length;i<len;i++){
        spreadsheet.deleteSheet(allSheets[i]);
      }
      
      // create
      allSheets[0].setName(sheetNames[0]);
      for(var sheetCounter = 1; sheetCounter < sheetNames.length; sheetCounter++) {

        var sheetName = sheetNames[sheetCounter];
        if(DEBUG == 1) Logger.log("sheet name: " + sheetName );
        spreadsheet.insertSheet(sheetName);
      }
    } else {
      // IF NOT NEW, MAKE SURE RIGHT SHEETS EXIST
      for(var sheetCounter = 0; sheetCounter < sheetNames.length; sheetCounter++) {
        var sheetName = sheetNames[sheetCounter];
        if(DEBUG == 1) Logger.log("checking if sheet with name exists: " + sheetName);
        try {
          var thisSheet = spreadsheet.getSheetByName(sheetName);
          if(!thisSheet) spreadsheet.insertSheet(sheetName);
        } catch (e) {
          Logger.log(e);
        }
      }
    }
    
    
    // ADD ACCOUNT MANAGERS
    if(accountManagers && accountManagers!=""){
      var accountManagersArray = accountManagers.replace(/\s/g, "").split(",");
      spreadsheet.addEditors(accountManagersArray);
    }
    
    // IF OVERWRITE, CLEAR SHEETS
    if(overWriteOldData) {
      for(var sheetCounter = 0; sheetCounter < sheetNames.length; sheetCounter++) {
        var sheetName = sheetNames[sheetCounter];
        if(DEBUG == 1) Logger.log("sheet name: " + sheetName);
        try {
          var thisSheet = spreadsheet.getSheetByName(sheetName);
          if(thisSheet) thisSheet.clear();
        } catch (e) {
          Logger.log(e);
        }
      }
    }
    return(destinationSpreadsheet);
  } 
