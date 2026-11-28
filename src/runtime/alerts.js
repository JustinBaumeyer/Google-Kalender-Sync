/**
 * Email warnings. An account's own scheduled runs warn that account. Problems that
 * stop runs altogether cannot report themselves, so every account also keeps a short
 * health record in the script properties, which the owner's daily check reads.
 */
var alertDelayMs = 3600000; // A problem must last this long before the first warning,
var alertReminderMs = 3 * 86400000; // and warnings about it repeat at most this often.

/** The address for warnings: the account's own, another one, or "" when switched off. */
function alertRecipient_(setting) {
  if (setting === true) return currentUserEmail_();
  return typeof setting == "string" ? setting.trim() : "";
}

function alertTime_(value) {
  return Utilities.formatDate(new Date(value), "Europe/Berlin", "dd.MM.yyyy 'um' HH:mm 'Uhr'");
}

function appUrl_() {
  try {
    return ScriptApp.getService().getUrl() || "";
  } catch (err) {
    return "";
  }
}

/** Sends one warning; a failed delivery is logged and never hides the run's own result. */
function sendAlertMail_(recipient, subject, lines) {
  var url = appUrl_();
  try {
    MailApp.sendEmail(
      recipient,
      "Dienstplan Sync: " + subject,
      lines.concat(url ? ["", "App öffnen: " + url] : []).join("\n")
    );
    return true;
  } catch (err) {
    log_("Could not send warning email: " + ((err && err.message) || err));
    return false;
  }
}

function readJson_(text, fallback) {
  try {
    var value = JSON.parse(text);
    return value && typeof value == "object" ? value : fallback;
  } catch (err) {
    return fallback;
  }
}

/**
 * Records the outcome of a finished run and sends the warnings that are due. A problem
 * is reported once it has lasted an hour (an expired roster access at once), then at
 * most every three days, and a short note follows once a run succeeds again. Runs
 * started by hand never send mail, because the person sees the result on screen.
 *
 * @param {?Error|string} error - Why the run failed, or null for a successful run
 * @param {boolean} scheduled - Whether a trigger started the run
 */
function trackRunOutcome_(error, scheduled) {
  try {
    var now = Date.now(),
      saved = userPrp.getProperties();
    var settings = readUserPreferences_(saved);
    var recipient = alertRecipient_(settings.errorNotificationEmail);
    var alert = readJson_(saved.alertState, {});
    if (!error) {
      if (alert.sentAt && recipient)
        sendAlertMail_(recipient, "Synchronisierung läuft wieder", [
          "Hallo,",
          "",
          "die automatische Synchronisierung deines Dienstplans läuft wieder.",
          "Das Problem bestand seit " + alertTime_(alert.since) + ".",
          "",
          "Es ist nichts weiter zu tun."
        ]);
      userPrp.deleteProperty("alertState");
      alert = {};
    } else {
      var message = String((error && error.message) || error).slice(0, 300);
      if (!alert.since) alert.since = now;
      alert.problem = message;
      alert.credential = !!(error && error.rosterAuthRejected);
      var due = alert.sentAt
        ? now - alert.sentAt >= alertReminderMs
        : alert.credential || now - alert.since >= alertDelayMs;
      if (scheduled && recipient && due && sendProblemMail_(recipient, alert)) alert.sentAt = now;
      userPrp.setProperty("alertState", JSON.stringify(alert));
    }
    writeAccountHealth_({
      at: new Date(now).toISOString(),
      ok: !error,
      since: alert.since ? new Date(alert.since).toISOString() : null,
      problem: alert.problem || null,
      credential: !!alert.credential,
      enabled: hasTrigger_("startSync"),
      interval: getValidTriggerFrequency_(settings.howFrequent),
      notify: recipient
    });
  } catch (err) {
    log_("Could not record the run outcome: " + ((err && err.message) || err));
  }
}

function sendProblemMail_(recipient, alert) {
  return sendAlertMail_(
    recipient,
    alert.credential ? "Dienstplan-Zugang abgelaufen" : "Synchronisierung gestört",
    [
      "Hallo,",
      "",
      "die automatische Synchronisierung deines Dienstplans funktioniert seit " +
        alertTime_(alert.since) +
        " nicht.",
      "",
      "Meldung: " + alert.problem,
      "",
      alert.credential
        ? "Bitte öffne die App und trage über „Neuen Token eintragen“ einen neuen Dienstplan-Token ein."
        : "Deine Kalendertermine bleiben unverändert, und jeder geplante Lauf versucht es erneut. Bleibt das Problem, findest du in der App unter „Fehlerdetails“ einen Bericht zur Weitergabe.",
      "",
      "Solange das Problem besteht, kommt höchstens alle drei Tage eine Erinnerung, und eine kurze Nachricht, sobald es wieder läuft.",
      "Abschalten lassen sich diese E-Mails in der App unter Einstellungen → „E-Mail bei Fehlern“."
    ]
  );
}

/** A run that only paused or continues later still shows that the triggers work. */
function noteRunActivity_() {
  writeAccountHealth_({ at: new Date().toISOString(), enabled: true });
}

/** The shared record the owner's daily check reads. Each account writes only its own. */
function writeAccountHealth_(patch) {
  try {
    var email = currentUserEmail_();
    if (!email) return;
    var key = "health:" + email;
    var record = readJson_(scriptPrp.getProperty(key), {});
    scriptPrp.setProperty(key, JSON.stringify(Object.assign(record, patch, { email: email })));
  } catch (err) {
    log_("Could not update the account health: " + ((err && err.message) || err));
  }
}

function deleteAccountHealth_() {
  var email = currentUserEmail_();
  if (!email) return;
  scriptPrp.deleteProperty("health:" + email);
  scriptPrp.deleteProperty("stallWarned:" + email);
}

/**
 * How an account is doing, judged from its record: "paused", "stalled" (enabled but no
 * run for three intervals, at least two hours), "failing", "ok" or "unknown".
 */
function accountHealthState_(record, now) {
  if (!record || !record.enabled) return { kind: "paused" };
  var interval = Math.max(Number(record.interval) || 15, 5) * 60000;
  var lastSign = Math.max(
    Date.parse(record.at || "") || 0,
    Date.parse(record.enabledAt || "") || 0
  );
  if (lastSign && now - lastSign > Math.max(3 * interval, 7200000))
    return { kind: "stalled", since: lastSign };
  if (record.ok === false && record.since) {
    var since = Date.parse(record.since);
    if (record.credential || now - since >= alertDelayMs)
      return { kind: "failing", since: since, problem: record.problem || "" };
  }
  return { kind: lastSign ? "ok" : "unknown", at: Date.parse(record.at || "") || null };
}

/** Health of every invited account, for the owner. Records of removed accounts are dropped. */
function readAccountHealth_(now) {
  var owner = appOwnerEmail_();
  var invited = new Set(
    [owner].concat(
      invitedEmails_().map(function (value) {
        return String(value).trim().toLowerCase();
      })
    )
  );
  var all = scriptPrp.getProperties(),
    result = {};
  Object.keys(all).forEach(function (key) {
    if (key.indexOf("health:") != 0) return;
    var record = readJson_(all[key], null);
    if (!record || !invited.has(record.email)) {
      scriptPrp.deleteProperty(key);
      return;
    }
    result[record.email] = Object.assign(accountHealthState_(record, now), {
      notify: record.notify || ""
    });
  });
  return result;
}

/**
 * The owner's daily check. Accounts whose runs stopped are warned directly, because
 * they cannot notice it themselves; the owner gets one summary of all accounts with
 * problems when that list changes, and at most every three days while it stays.
 */
function reviewAccountHealth_() {
  var now = Date.now(),
    health = readAccountHealth_(now),
    problems = [];
  Object.keys(health)
    .sort()
    .forEach(function (email) {
      var state = health[email],
        warnedKey = "stallWarned:" + email;
      if (state.kind != "stalled") {
        scriptPrp.deleteProperty(warnedKey);
        if (state.kind == "failing")
          problems.push({
            email: email,
            kind: state.kind,
            since: state.since,
            text: "Fehler seit " + alertTime_(state.since) + ": " + state.problem
          });
        return;
      }
      problems.push({
        email: email,
        kind: state.kind,
        since: state.since,
        text:
          "kein automatischer Lauf seit " +
          alertTime_(state.since) +
          " – ist das Skriptprojekt mit diesem Konto als Betrachter geteilt? Ohne diese Freigabe lehnt Google automatische Läufe ab."
      });
      var warned = Date.parse(scriptPrp.getProperty(warnedKey) || "") || 0;
      if (state.notify && now - warned >= alertReminderMs) {
        var sent = sendAlertMail_(state.notify, "Automatische Synchronisierung angehalten", [
          "Hallo,",
          "",
          "für dein Konto hat seit " +
            alertTime_(state.since) +
            " kein automatischer Lauf mehr stattgefunden.",
          "Google schaltet die automatischen Läufe zum Beispiel ab, wenn eine Freigabe abgelaufen ist, und erneute Zustimmung schaltet sie nicht wieder ein.",
          "",
          "Bitte öffne die App einmal. Dabei wird die automatische Synchronisierung neu eingerichtet; falls Google nach deiner Zustimmung fragt, bestätige sie. Einstellungen und Termine bleiben erhalten.",
          "Hilft das nicht, muss der App-Inhaber das Skriptprojekt für dein Konto freigeben. Er erhält dazu ebenfalls eine Nachricht.",
          "",
          "Diese Nachricht kommt aus der täglichen Prüfung des App-Inhabers.",
          "Abschalten lassen sich diese E-Mails in der App unter Einstellungen → „E-Mail bei Fehlern“."
        ]);
        if (sent) scriptPrp.setProperty(warnedKey, new Date(now).toISOString());
      }
    });
  var signature = JSON.stringify(
    problems.map(function (problem) {
      return [problem.email, problem.kind, problem.since];
    })
  );
  var last = readJson_(scriptPrp.getProperty("healthSummary"), {});
  var owner = appOwnerEmail_();
  var due = problems.length
    ? signature != last.signature || now - (last.at || 0) >= alertReminderMs
    : !!last.signature && last.signature != "[]";
  if (!due || !owner) return;
  var sent = sendAlertMail_(
    owner,
    problems.length
      ? problems.length + (problems.length == 1 ? " Konto braucht" : " Konten brauchen") + " Hilfe"
      : "Alle Konten laufen wieder",
    problems.length
      ? ["Diese Konten synchronisieren gerade nicht:", ""]
          .concat(
            problems.map(function (problem) {
              return "• " + problem.email + " – " + problem.text;
            })
          )
          .concat([
            "",
            "Wer Warnungen nicht abgeschaltet hat, wurde selbst benachrichtigt. Den Stand aller Konten zeigt die App unter „Zugriff verwalten“."
          ])
      : ["Alle zuletzt gemeldeten Konten synchronisieren wieder."]
  );
  if (sent)
    scriptPrp.setProperty("healthSummary", JSON.stringify({ signature: signature, at: now }));
}
