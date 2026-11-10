function refreshRosterToken_() {
  rosterUserToken = JSON.parse(rosterFetch_("auth/refreshToken", "GET")).data;
  userPrp.setProperty("rosterUserToken", rosterUserToken);
}
