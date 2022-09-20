// Automate weekly PPC report

var NARRATIVE_SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1lX3klfgnfpnwrfQG4hVYwbqUFqlTrCR_-gYM7c4Cecs/'
var SHEET_NAME = 'automation';

function replaceMetric(pages) {
  var text_vals = getAutoValue(); 
  for (var i = 0; i < pages.length; i++) {
     var page = pages[i];
     replaceTextByAutoVal(page, text_vals);
  }

}

function replaceSelectedSlides() {
  var selection = SlidesApp.getActivePresentation().getSelection();  
  var selectionType = selection.getSelectionType();
  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var pageRange = selection.getPageRange();
    var pages = pageRange.getPages();
    replaceMetric(pages);
  }
}

function replaceTextByAutoVal(slide, text_vals) {
  slide.replaceAllText("{date}", text_vals[0][0]);
  slide.replaceAllText("{account1_search_conv_raw}", text_vals[1][0]);
  slide.replaceAllText("{account1_search_conv_mom}", text_vals[2][0]);
  slide.replaceAllText("{account1_search_cpa_raw}", text_vals[3][0]);
  slide.replaceAllText("{account1_search_cpa_mom}", text_vals[4][0]);
  slide.replaceAllText("{account1_search_cpc_raw}", text_vals[5][0]);
  slide.replaceAllText("{account1_search_cpc_mom}", text_vals[6][0]);
  slide.replaceAllText("{account1_display_conv_raw}", text_vals[7][0]);
  slide.replaceAllText("{account1_display_conv_mom}", text_vals[8][0]);
  slide.replaceAllText("{account1_display_cpa_raw}", text_vals[9][0]);
  slide.replaceAllText("{account1_display_cpa_mom}", text_vals[10][0]);
  slide.replaceAllText("{account1_display_cpc_raw}", text_vals[11][0]);
  slide.replaceAllText("{account1_display_cpc_mom}", text_vals[12][0]);
  slide.replaceAllText("{account1_fb_leads_raw}", text_vals[13][0]);
  slide.replaceAllText("{account1_fb_leads_mom}", text_vals[14][0]);
  slide.replaceAllText("{account1_fb_cpa_raw}", text_vals[15][0]);
  slide.replaceAllText("{account1_fb_cpa_mom}", text_vals[16][0]);
  slide.replaceAllText("{account2_search_conv_raw}", text_vals[17][0]);
  slide.replaceAllText("{account2_search_conv_mom}", text_vals[18][0]);
  slide.replaceAllText("{account2_search_cpa_raw}", text_vals[19][0]);
  slide.replaceAllText("{account2_search_cpa_mom}", text_vals[20][0]);
  slide.replaceAllText("{account2_search_cpc_raw}", text_vals[21][0]);
  slide.replaceAllText("{account2_search_cpc_mom}", text_vals[22][0]);
  slide.replaceAllText("{account2_display_conv_raw}", text_vals[23][0]);
  slide.replaceAllText("{account2_display_conv_mom}", text_vals[24][0]);
  slide.replaceAllText("{account2_display_cpa_raw}", text_vals[25][0]);
  slide.replaceAllText("{account2_display_cpa_mom}", text_vals[26][0]);
  slide.replaceAllText("{account2_display_cpc_raw}", text_vals[27][0]);
  slide.replaceAllText("{account2_display_cpc_mom}", text_vals[28][0]);
  slide.replaceAllText("{account2_video_conv_raw}", text_vals[29][0]);
  slide.replaceAllText("{account2_video_conv_mom}", text_vals[30][0]);
  slide.replaceAllText("{account2_video_cpa_raw}", text_vals[31][0]);
  slide.replaceAllText("{account2_video_cpa_mom}", text_vals[32][0]);
  slide.replaceAllText("{account2_video_cpv_raw}", text_vals[33][0]);
  slide.replaceAllText("{account2_video_cpv_mom}", text_vals[34][0]);
  slide.replaceAllText("{account3_search_conv_raw}", text_vals[35][0]);
  slide.replaceAllText("{account3_search_conv_mom}", text_vals[36][0]);
  slide.replaceAllText("{account3_search_cpa_raw}", text_vals[37][0]);  
  slide.replaceAllText("{account3_search_cpa_mom}", text_vals[38][0]);
  slide.replaceAllText("{account3_search_cpc_raw}", text_vals[39][0]);
  slide.replaceAllText("{account3_search_cpc_mom}", text_vals[40][0]);
  Logger.log("Replacement done");
}

function getAutoValue() {
  var sheet = SpreadsheetApp.openByUrl(NARRATIVE_SHEETS_URL).getSheetByName(SHEET_NAME);
  var range = sheet.getRange("D2:D50");
  var values = range.getValues();
  return values;
}

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Replace Metrics', 'replaceSelectedSlides')
      .addToUi();
}
