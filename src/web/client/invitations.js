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
async function updateInvitations(emails) {
  busy(true);
  try {
    renderInvitations(await call("saveInvitations", emails));
    $("inviteEmail").value = "";
    $("adminMessage").textContent = "Zugriff gespeichert. Den App-Link bitte selbst weitergeben.";
  } catch (error) {
    $("adminMessage").textContent = error.message || String(error);
  } finally {
    busy(false);
  }
}
async function openInvitations() {
  busy(true);
  try {
    renderInvitations(await call("getInvitations"));
    $("adminPanel").scrollIntoView({ behavior: "smooth" });
  } catch (error) {
    message(error.message, true);
  } finally {
    busy(false);
  }
}
