function generateICalEntry(uid, start,end,summary,description) {
  return "BEGIN:VEVENT\nUID:" + uid + "\nSEQUENCE:0\nDTSTAMP:" + new Date().toISOString() + "\nDTSTART:" + start + "\nDTEND:" + end + "\nSUMMARY:" + summary + "\nDESCRIPTION:"+(description)+"\nEND:VEVENT\n";
}

function parseShiftToCal(shift) {
    var ret = "";
    const dienste = new Map();
    shift.data.rosterDetails.entries.forEach(dienst => {
        if (!rosterIgnoreList.includes(dienst.shortName)) {
            var start = new Date(dienst.from)
            var end = new Date(dienst.to)
            var mapName = dienst.shortName;
            if(dienste.has(mapName)) {
              var d = dienste.get(dienst.shortName)
              if (dienst.start == d.end || dienst.end == d.start) {
                 mapName = dienst.shortName+dienst.start;
              } else {
                if(d.start > start) d.start = start;
                if(d.end < end) d.end = end;
                dienste.set(mapName,d);
              }
            }
        }
        if(!dienste.has(mapName)) {                     
          dienste.set(mapName,{"start":start,"end":end,"nameWorkplace":dienst.nameWorkplace,"shortName": dienst.shortName});            
        }
    })
    var summary = [];
    dienste.forEach(dienst => {
      var start = Utilities.formatDate(new Date(dienst.start),"GMT","yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[\-,\:]/g, '')
      var end = Utilities.formatDate(new Date(dienst.end),"GMT","yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[\-,\:]/g, '')
      ret += generateICalEntry(dienst.shortName + start+end, start,end,dienst.shortName + " | " + dienst.nameWorkplace,"")
      summary.push(dienst.shortName);
    })
    return {"ical": ret, "list": summary};
}

function refreshRosterToken() {
    var urlResponse = UrlFetchApp.fetch(rosterUrl+"auth/refreshToken", {
        'validateHttpsCertificates': false,
        'muteHttpExceptions': true,
        "headers": {
            "authorization": "Bearer " + rosterUserToken,
            "content-type": "application/json",
        },
        "method": "GET"
    });
    if (urlResponse.getResponseCode() == 200) {
        rosterUserToken = JSON.parse(urlResponse.getContentText()).data;
        scriptPrp.setProperty('rosterUserToken', rosterUserToken);
    } else { //Throw here to make callWithBackoff run again
        throw "Error: Encountered HTTP error " + urlResponse.getContentText() + " when accessing refreshToken";
    }
}

function getRosterStartDate() {
    return callWithBackoff(function() {
      var urlResponse = UrlFetchApp.fetch(rosterUrl+"app/initial-preload", {
          'validateHttpsCertificates': false,
          'muteHttpExceptions': true,
          "headers": {
              "authorization": "Bearer " + rosterUserToken,
              "content-type": "application/json",
          },
          "method": "POST"
      });
      if (urlResponse.getResponseCode() == 200) {
          var jsonResponse = JSON.parse(urlResponse.getContentText())
          var date = jsonResponse.contracts.data[0].eingestellt.split(".");
          return new Date(date[2], date[1] - 1, date[0])
      } else { //Throw here to make callWithBackoff run again
          throw "Error: Encountered HTTP error " + urlResponse.getContentText() + " when accessing initial-preload";
      }
    },defaultMaxRetries);
}

var globalStartDate = null;
var globalEndDate = null;

function generateRosterPayload() {
    var payload = "["
    var startDate = new Date();
    var endDate = new Date();
    endDate.setMonth(endDate.getMonth() + 2);
    if (addRosterSinceStart) {
        startDate = getRosterStartDate();
    }
    if(addYearSummary) {
      startDate = new Date("1.1."+summaryYear)
      endDate = new Date(summaryYear, 11, 31)
    }
    globalStartDate = startDate.toISOString();
    globalEndDate = endDate.toISOString();
    while (endDate - startDate > 0) {
        startDate.setDate(1);
        payload += "{\"employeeId\":" + rosterUserId + ",\"begin\":\"" + startDate.toISOString() + "\",\"end\":\"";
        startDate.setMonth(startDate.getMonth() + 1, 0);
        payload += startDate.toISOString() + "\",\"rosterViewMode\":4},"
        startDate.setDate(startDate.getDate() + 1)
    }
    payload.slice(0, -1);
    payload += "]"

    return payload;
}

function getRosterICal() {
    callWithBackoff(function() {
        refreshRosterToken();
        return;
    },defaultMaxRetries);
    return callWithBackoff(function() {
        var urlResponse = UrlFetchApp.fetch(rosterUrl+"rosters/preload", {
            'validateHttpsCertificates': false,
            'muteHttpExceptions': false,
            "headers": {
                "authorization": "Bearer " + rosterUserToken,
                "content-type": "application/json",
            },
            "payload": generateRosterPayload(),
            "method": "POST"
        });
      if (urlResponse.getResponseCode() == 200) {
            var jsonResponse = JSON.parse(urlResponse.getContentText())
            var icsContent = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//GoogleKalenderSync/EN\nMETHOD:REQUEST\nNAME:Roster\nX-WR-CALNAME:Roster\n";
            
            var dienstCount = new Map();
            jsonResponse.forEach(month => {
                month.rosterDetails.forEach(shift => {
                  var data = parseShiftToCal(shift)
                    icsContent += data.ical
                    if (addYearSummary) {
                        data.list.forEach(d => {
                            if(dienstCount.has(d)) {
                                dienstCount.set(d, dienstCount.get(d)+1)
                            } else {
                                dienstCount.set(d, 1)
                            }
                        })
                    }
                })
                if(addRosterRequests) {
                  month.rosterUrlaubEinsatzwunsch.data.data[0].item2.data.forEach(day => {
                    if(day.einsatzwunsch) {
                      var date = Utilities.formatDate(new Date(day.day.date),"GMT","yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[\-,\:]/g, '')
                      var summary = "Einsatzwunsch: " + (function(ew){
                        switch(ew){
                          case 1: return "Frei";
                          case 2: return "Beliebiger Dienst";
                          case 3: return "Bestimmter Dienst";
                          case 4: return "Dienst ausschließen";
                          default: return "Frei";
                        }})(day.einsatzwunschWunsch);
                      summary += " " + day.shortName;
                      
                      icsContent += generateICalEntry(day.einsatzwunschWunsch+"-"+day.shortName+"-"+date,date,date,summary,day.kommentar)
                    }
                  })
                }
            });

            if (addYearSummary) {
              Logger.log("Year: " + summaryYear);
              var loggerText = ""
              dienstCount = new Map([...dienstCount.entries()].sort());
              dienstCount.forEach((key, value) => {
                  loggerText += value + ": " + key + ","
              })
              Logger.log(loggerText.slice(0,-1))
            }
            if(updateAvailable== "true") {
              var date = Utilities.formatDate(new Date(),"GMT","yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[\-,\:]/g, '')
              icsContent += generateICalEntry(date+"-updateNotification",date,date,"Dienstplan Update verfügbar","Es ist ein Update für das Dienstplanprogramm verfügbar. Bitte den Updateanweisungen in der Gruppe folgen.");
              Logger.log("update updateAvailable " + updateAvailable)
            }

            icsContent += "END:VCALENDAR"
            return icsContent;
        } else { //Throw here to make callWithBackoff run again
            throw "Error: Encountered HTTP error " + urlResponse.getContentText() + " when accessing " + url;
        }
    }, defaultMaxRetries);
}