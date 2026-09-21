# Call Tree Dashboard V2: GitHub Pages + Microsoft Graph

Static website with no Node.js build step. It signs users in with Microsoft Entra ID and reads the SharePoint List `Member` through Microsoft Graph.

## SharePoint source
- Site: `https://singhaestate.sharepoint.com/sites/ITProject`
- List: `Member`
- Browser view: `https://singhaestate.sharepoint.com/sites/ITProject/Lists/Member/AllItems.aspx`

## 1. Microsoft Entra App Registration
1. Create an app registration for this dashboard.
2. Under **Authentication**, add platform **Single-page application**.
3. Add the exact Redirect URI: `https://YOUR_GITHUB_USERNAME.github.io/calltree-dashboard/`
4. Under Microsoft Graph delegated permissions, add `User.Read` and `Sites.Read.All`.
5. Complete admin consent according to company policy.
6. Keep the application client secret empty. A static browser app must not contain a client secret.

## 2. Edit config.js
Replace:
- `REPLACE_WITH_ENTRA_APPLICATION_CLIENT_ID`
- `REPLACE_WITH_GITHUB_USERNAME`

The tenant ID, SharePoint hostname, site path, and list name are already configured.

## 3. Upload to GitHub
Upload the contents of this ZIP to the root of a repository named `calltree-dashboard`. Do not upload the containing folder as an extra directory.

## 4. Enable GitHub Pages
In the repository, open **Settings > Pages** and set **Source** to **GitHub Actions**. Push to `main`. The included workflow publishes the site.

## 5. Open
`https://YOUR_GITHUB_USERNAME.github.io/calltree-dashboard/`

## Expected SharePoint internal field names
`Mode`, `Member`, `EMail`, `ClickDateTime`, `Created`, `iMsg`, `ResponseSafe`, `HelpNote`, `RefID`, `DrillChoice`, and `DrillResponse`.

If a Graph error reports an invalid field, confirm the SharePoint internal name and update the `fields` constant in `app.js`.

## Security
- The website contains no secret.
- Access tokens are kept in session storage by MSAL.
- Users still need permission to the SharePoint site/list.
- `Sites.Read.All` is broad delegated access and requires organizational review/consent.
