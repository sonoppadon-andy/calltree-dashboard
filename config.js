// Replace CLIENT_ID with the Application (client) ID from Microsoft Entra App Registration.
// This file contains no client secret. Do not add a secret to a browser application.
window.CALLTREE_CONFIG = {
  clientId: "be89b11f-ca7d-4181-8a0a-20e8199ecc05",
  tenantId: "1f20664a-d62c-485d-8200-352c44549547",
  redirectUri: "https://sonoppadon-andy.github.io/calltree-dashboard",
  sharePointHost: "singhaestate.sharepoint.com",
  sitePath: "/sites/ITProject",
  listName: "Member",
  graphScopes: ["User.Read", "Sites.Read.All"]
};
