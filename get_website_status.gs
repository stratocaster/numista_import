function getStatusCode(url){
   var response = UrlFetchApp.fetch(url);
   return response.getResponseCode();
}
