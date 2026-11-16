function renderEventList(container, events) {
  container.replaceChildren();
  if (!events.length) {
    container.append(element("li", "Keine kommenden Termine in der Vorschau.", "muted"));
    return;
  }
  events.forEach((event) => {
    const row = element("li"),
      start = event.allDay ? new Date(event.start + "T12:00:00") : new Date(event.start);
    let when;
    if (event.allDay) {
      const lastDay = new Date(Date.parse(event.end + "T00:00:00Z") - 86400000).toLocaleDateString(
        "de-DE",
        { timeZone: "UTC" }
      );
      const firstDay = start.toLocaleDateString("de-DE");
      when = firstDay + (lastDay !== firstDay ? " – " + lastDay : "") + " · ganztägig";
    } else {
      const end = new Date(event.end),
        timeZone = current.settings.rosterTimeZone;
      const sameDay =
        start.toLocaleDateString("de-DE", { timeZone }) ===
        end.toLocaleDateString("de-DE", { timeZone });
      when =
        start.toLocaleString("de-DE", { timeZone }) +
        " – " +
        (sameDay
          ? end.toLocaleTimeString("de-DE", { timeZone, hour: "2-digit", minute: "2-digit" })
          : end.toLocaleString("de-DE", { timeZone }));
    }
    row.append(element("strong", event.title), element("small", when + " · " + event.calendar));
    container.append(row);
  });
}
async function diagnostic(kind) {
  let payload;
  try {
    payload =
      kind === "preview"
        ? collect()
        : Object.assign(
            {
              settings: {
                rosterUrl: readers.get("rosterUrl")(),
                addRosterToCal: readers.get("addRosterToCal")()
              }
            },
            collectRosterCredentials(readers.get("addRosterToCal")())
          );
  } catch (error) {
    message(error.message, true);
    return;
  }
  message(kind === "preview" ? "Vorschau wird geladen …" : "Verbindung wird geprüft …");
  await whileBusy(
    async () => {
      const result = await call(
        kind === "preview" ? "previewSynchronization" : "testRosterConnection",
        payload
      );
      if (kind === "preview") {
        $("previewPanel").hidden = false;
        previewChecked = true;
        if (dashboardState) renderSetupGuide(dashboardState);
        renderEventList($("previewEvents"), result.events);
        $("previewErrors").replaceChildren();
        result.errors.forEach((error) =>
          $("previewErrors").append(element("p", error.calendar + ": " + error.message))
        );
        $("previewNote").textContent = result.note;
        message(
          result.errors.length
            ? "Einige Quellen konnten nicht geladen werden."
            : "Vorschau geladen. Im Kalender wurde nichts geändert."
        );
      } else {
        if (!dirty) render(await call("getAppState"));
        $("connectionResult").textContent = result.message;
        message(result.message);
      }
    },
    (text) => {
      message(text, true);
      if (kind === "test") $("connectionResult").textContent = text;
    }
  );
}

function renderInvitations(result) {
  invitations = result;
  $("adminPanel").hidden = false;
  $("ownerInfo").textContent = "App-Inhaber: " + result.owner + " (Zugriff bleibt immer erhalten)";
  $("inviteList").replaceChildren();
  result.emails
    .filter((email) => email !== result.owner)
    .forEach((email) => {
      const row = element("li"),
        remove = element("button", "Zugriff entziehen", "remove");
      remove.type = "button";
      remove.setAttribute("aria-label", "Zugriff für " + email + " entziehen");
      row.append(element("span", email), remove);
      $("inviteList").append(row);
      remove.onclick = () => updateInvitations(result.emails.filter((value) => value !== email));
    });
}
function updateInvitations(emails) {
  return whileBusy(async () => {
    renderInvitations(await call("saveInvitations", emails));
    $("inviteEmail").value = "";
    $("adminMessage").textContent = "Zugriff gespeichert. Den App-Link bitte selbst weitergeben.";
  }, showIn("adminMessage"));
}
function openInvitations() {
  return whileBusy(
    async () => {
      renderInvitations(await call("getInvitations"));
      $("adminPanel").scrollIntoView({ behavior: "smooth" });
    },
    (text) => message(text, true)
  );
}

function openMigration() {
  $("migrationPanel").hidden = false;
  $("migrationPanel").scrollIntoView({ behavior: "smooth" });
}
function copyMigration() {
  return whileBusy(async () => {
    const copied = await copyText(await call("getMigrationSnippet"), $("migrationSnippet"));
    $("migrationMessage").textContent = copied
      ? "Kopiert. In eine neue Datei der alten Kopie einfügen und prepareSharedAppMigration ausführen."
      : "Den markierten Text kopieren und unverändert in eine neue Datei der alten Kopie einfügen.";
  }, showIn("migrationMessage"));
}
function importMigration() {
  return whileBusy(async () => {
    const state = await call("importLegacyCopy", {
      code: $("migrationCode").value,
      replaceExisting: $("replaceMigration").checked
    });
    $("migrationCode").value = "";
    $("migrationPanel").hidden = true;
    render(state);
    showSettings(true);
    message(
      "Einstellungen und Zugangsdaten übernommen. Prüfe Verbindung, Zielkalender und Vorschau. Danach kannst du den neuen Sync aktivieren."
    );
  }, showIn("migrationMessage"));
}
