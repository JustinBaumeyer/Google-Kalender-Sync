/** Paste this entire file into a NEW file in the OLD private script copy.
 * Run prepareSharedAppMigration. No values or existing files need editing.
 * The generated code contains credentials and is never written to logs.
 */
function prepareSharedAppMigration() {
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!email || Session.getActiveUser().getEmail().trim().toLowerCase() != email)
    throw new Error("Bitte direkt im Skripteditor mit deinem Google-Konto ausführen.");
  var project = PropertiesService.getScriptProperties(), personal = PropertiesService.getUserProperties();
  var owner = project.getProperty("appOwnerEmail");
  if (owner && owner != email) throw new Error("Bitte deine eigene bisherige Skriptkopie verwenden.");
  var scriptLock = LockService.getScriptLock(), userLock = LockService.getUserLock(), userLocked = false;
  if (!scriptLock.tryLock(0)) throw new Error("Die bisherige Synchronisierung läuft noch. Bitte danach erneut versuchen.");
  try {
    userLocked = userLock.tryLock(0);
    if (!userLocked) throw new Error("Die bisherige Synchronisierung läuft noch. Bitte danach erneut versuchen.");
    var saved = personal.getProperties(), settings = saved.preferences ? JSON.parse(saved.preferences) : {};
    var names = ["sourceCalendars","howFrequent","onlyFutureEvents","addEventsToCalendar","modifyExistingEvents","removeEventsFromCalendar","removePastEventsFromCalendar","addCalToTitle","defaultAllDayReminder","showSkipMessage","overrideVisibility","errorNotificationEmail","consolidateEvents","consolidateMaxGapMinutes","ignoreEvents","addRosterToCal","addRosterSinceStart","rosterTimeZone","addRosterRequests","addAbsences","oncallAsFree","addTeamPartner","rosterPlanningGroups","addRelief","addYearSummary","summaryYear","rosterUrl","rosterIgnoreList"];
    if (!saved.preferences) names.forEach(function(name) { if (globalThis[name] !== undefined) settings[name] = globalThis[name]; });
    if (settings.consolidateEvents === undefined && typeof consolidateShifts !== "undefined") settings.consolidateEvents = !!consolidateShifts;
    if (!Array.isArray(settings.sourceCalendars)) throw new Error("Die bisherigen Kalenderquellen konnten nicht gelesen werden.");
    var calendars = [], pageToken;
    do {
      var options = {showHidden:true,minAccessRole:"writer",maxResults:250};
      if (pageToken) options.pageToken = pageToken;
      var page = Calendar.CalendarList.list(options);
      calendars = calendars.concat((page.items || []).filter(function(cal) { return ["owner","writer"].indexOf(cal.accessRole) != -1; }));
      pageToken = page.nextPageToken;
    } while (pageToken);
    settings.sourceCalendars = settings.sourceCalendars.map(function(row) {
      var result = row.slice();
      if (!result[3]) {
        var matches = calendars.filter(function(cal) { return (cal.summaryOverride || cal.summary) == row[1]; });
        if (matches.length == 1) { if (result.length < 3) result.push(null); result[3] = matches[0].id; }
      }
      return result;
    });
    var payload = {format:"dienstplan-migration",version:1,email:email,sourceScriptId:ScriptApp.getScriptId(),
      expiresAt:Date.now()+86400000,oldSyncPaused:true,settings:settings,
      rosterUserId:saved.rosterUserId || project.getProperty("rosterUserId") || "",
      rosterUserToken:saved.rosterUserToken || project.getProperty("rosterUserToken") || ""};
    var text = JSON.stringify(payload,function(key,value) {
      return value instanceof RegExp ? {$regex:value.source,$flags:value.flags} : value;
    });
    var code = "DPSYNC1." + Utilities.base64EncodeWebSafe(text,Utilities.Charset.UTF_8);
    if (Utilities.newBlob(code).getBytes().length > 9000)
      throw new Error("Die Einstellungen sind zu groß für einen Umzugscode. Bitte den App-Inhaber kontaktieren; es wurde nichts pausiert.");
    project.setProperty("DienstplanMigration","Wird vorbereitet – noch nicht importieren");
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      if (["startSync","continueSync_","install","main","checkForUpdates"].indexOf(trigger.getHandlerFunction()) != -1)
        ScriptApp.deleteTrigger(trigger);
    });
    project.setProperty("DienstplanMigration",code);
    Logger.log("Die bisherige Synchronisierung ist pausiert. Kopiere den Wert DienstplanMigration unter Projekteinstellungen → Skripteigenschaften in die gemeinsame App. Der Code enthält Zugangsdaten, ist 24 Stunden gültig und gehört nur in die gemeinsame App. Nach erfolgreichem Umzug die Eigenschaft und die Hilfsdatei löschen.");
  } finally {
    if (userLocked) userLock.releaseLock();
    scriptLock.releaseLock();
  }
}
