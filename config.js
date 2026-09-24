// Replace CLIENT_ID with the Application (client) ID from Microsoft Entra App Registration.
// This file contains no client secret. Do not add a secret to a browser application.
window.CALLTREE_CONFIG = {
  clientId: "be89b11f-ca7d-4181-8a0a-20e8199ecc05",
  tenantId: "1f20664a-d62c-485d-8200-352c44549547",
  redirectUri: "https://sonoppadon-andy.github.io/calltree-dashboard",
  sharePointHost: "singhaestate.sharepoint.com",
  sitePath: "/sites/ITProject",
  listName: "Member",
  graphScopes: ["User.Read", "Sites.Read.All"],
  // Master employee list ("Phone Book") used to compute who has NOT responded yet.
  // Site: https://singhaestate.sharepoint.com/sites/snet/Lists/Phone%20Book/
  // Field names below are the SharePoint INTERNAL field names (confirmed via each
  // column's FldEdit.aspx URL 2026-09-23) — NOT the Display Names shown in the UI.
  // Do not rename these to the Display Names (Email Address/BU/Division/Job Title);
  // Graph API needs the internal "field_NN" names in $select or it silently returns blank.
  phoneBook: {
    sitePath: "/sites/snet",
    listName: "Phone Book",
    fields: { email: "field_15", bu: "field_16", division: "field_14", jobTitle: "field_10", department: "field_12" }
  }
};
