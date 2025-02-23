function checkUrlsStatusAnchorAndRel() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var urls = sheet.getRange("A2:A" + sheet.getLastRow()).getValues(); 
    var checkUrls = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
    var anchorTexts = sheet.getRange("E2:E" + sheet.getLastRow()).getValues(); 
    var relTypes = sheet.getRange("G2:G" + sheet.getLastRow()).getValues(); 
  
    var statusResults = [];
    var checkResults = [];
    var anchorResults = [];
    var relResults = [];
  
    for (var i = 0; i < urls.length; i++) {
      var pageUrl = urls[i][0]; 
      var checkUrl = checkUrls[i][0]; 
      var anchorText = anchorTexts[i][0];
      var expectedRel = relTypes[i][0];
  
      var pageStatus = "Error";
      var checkStatus = "Error";
      var anchorCheck = "Error";
      var relCheck = "Error";
  
      if (pageUrl) {
        try {
          var response = UrlFetchApp.fetch(pageUrl, {muteHttpExceptions: true});
          pageStatus = response.getResponseCode(); 
          var htmlContent = response.getContentText(); 
          if (checkUrl && htmlContent.includes(checkUrl)) {
            try {
              var checkResponse = UrlFetchApp.fetch(checkUrl, {muteHttpExceptions: true});
              checkStatus = checkResponse.getResponseCode(); 
  
              
              var anchorRegex = new RegExp('<a[^>]+href=["\']' + checkUrl.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + '["\'][^>]*>(.*?)</a>', 'i');
              var match = htmlContent.match(anchorRegex);
              
              if (match && match[1].trim() === anchorText.trim()) {
                anchorCheck = "200"; 
              }
  
         
              var relRegex = new RegExp('<a[^>]+href=["\']' + checkUrl.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + '["\'][^>]*rel=["\']([^"\']+)["\']', 'i');
              var relMatch = htmlContent.match(relRegex);
              
              if (relMatch) {
                var actualRel = relMatch[1].toLowerCase();
                var isFollow = !actualRel.includes("nofollow"); 
  
                if ((expectedRel.toLowerCase() === "follow" && isFollow) || 
                    (expectedRel.toLowerCase() === "nofollow" && !isFollow)) {
                  relCheck = "200"; 
                }
              } else if (expectedRel.toLowerCase() === "follow") {
                relCheck = "200"; 
              }
            } catch (err) {
              checkStatus = "Error";
            }
          }
        } catch (e) {
          pageStatus = "Error";
        }
      }
  
      statusResults.push([pageStatus]);
      checkResults.push([checkStatus]);
      anchorResults.push([anchorCheck]);
      relResults.push([relCheck]);
    }
  
    sheet.getRange("B2:B" + (statusResults.length + 1)).setValues(statusResults); 
    sheet.getRange("D2:D" + (checkResults.length + 1)).setValues(checkResults); 
    sheet.getRange("F2:F" + (anchorResults.length + 1)).setValues(anchorResults); 
    sheet.getRange("H2:H" + (relResults.length + 1)).setValues(relResults); 
    // Made by Aref Naghizadeh
  }
  