var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(
  Number(scriptPrp.getProperty("availableVersion") || version) > version
);
