# Innova Attachment List Add-in

This repository contains a minimal Outlook add-in for the New Outlook compose experience.

The add-in adds a ribbon button named `Insert Attachment List`. When clicked while composing a message, it:

- reads the current message's attachments with `Office.context.mailbox.item.getAttachmentsAsync`
- ignores inline/signature attachments
- inserts a formatted list at the current cursor location using `body.setSelectedDataAsync`

The inserted output matches the old VBA behavior closely:

```text
Encl. (2 files)
Proposal.pdf
Schedule.xlsx
```

## Files

- `manifest.xml`: add-in manifest for sideloading
- `appPackage/manifest.json`: unified Microsoft 365 manifest for Teams/Microsoft 365 upload
- `build-github-pages-release.ps1`: generates a Pages-ready manifest and Teams zip that point at GitHub Pages
- `build-pages-site.ps1`: creates the static site payload published to GitHub Pages
- `host/InnovaOutlook.Host.csproj`: local HTTPS static-file host
- `host/Program.cs`: host startup and static-file configuration
- `src/commands/commands.html`: function file loaded by Outlook
- `src/commands/commands.js`: ribbon command logic
- `build-app-package.ps1`: creates the Teams-uploadable zip package

## Local hosting

Office add-ins must be served over HTTPS. This repository now defaults to the GitHub Pages host at `https://driegonunez.github.io/InnovaOutlook`.

This repo still includes a small ASP.NET Core host in `host/` for local troubleshooting, but the default manifests now point to GitHub Pages instead of localhost.

Start it from the repository root:

```powershell
dotnet dev-certs https --trust
.\start-host.ps1
```

If you prefer not to use the helper script, this works too:

```powershell
dotnet run --project .\host\InnovaOutlook.Host.csproj
```

## GitHub Pages deployment

If you want the add-in to run without a local server, host the static files on GitHub Pages and then build a package that points to that Pages URL.

This repo is preconfigured to assume:

- `https://driegonunez.github.io/InnovaOutlook`

If your final GitHub Pages URL is different, pass `-BaseUrl` to the release script.

Prepare the static site payload:

```powershell
.\build-pages-site.ps1
```

Create a GitHub Pages release bundle:

```powershell
.\build-github-pages-release.ps1
```

That generates:

- `dist\github-pages\site\` for GitHub Pages hosting
- `dist\github-pages\manifest.xml` for Outlook-only sideload
- `dist\github-pages\InnovaAttachmentList-M365.zip` for Teams/Microsoft 365 upload

The repo also includes [deploy-pages.yml](C:\Users\dmendoza\Documents\GitHub\InnovaOutlook\.github\workflows\deploy-pages.yml), which publishes `dist/site` to GitHub Pages when you push to `main`.

## Teams / Microsoft 365 package

This repo also includes a unified manifest in `appPackage/manifest.json` so the add-in can be uploaded as a custom Microsoft 365 app package.

Build the zip package from the repository root:

```powershell
.\build-app-package.ps1
```

That creates:

- `dist/InnovaAttachmentList-M365.zip`

Upload that zip in Teams or the Microsoft 365 app upload flow.

For a GitHub Pages-backed deployment, upload `dist\github-pages\InnovaAttachmentList-M365.zip` instead.

Typical test flow:

1. Start the local host with `.\start-host.ps1`
2. Build the package with `.\build-app-package.ps1`
3. In Teams, choose `Apps` > `Manage your apps` > `Upload a custom app`
4. Select `dist/InnovaAttachmentList-M365.zip`
5. Open New Outlook and start a new message
6. Look for `Innova Tools` > `Insert Attachment List`

If you previously installed the XML manifest directly in Outlook, remove that older custom add-in first if you see a duplicate entry.

## Outlook-only sideload

Then verify:

- `https://localhost:3000/manifest.xml`
- `https://localhost:3000/src/commands/commands.html`

After those load without certificate warnings, sideload `manifest.xml` in New Outlook.

If you host from a different origin, update the localhost URLs in:

- `manifest.xml`
- `appPackage/manifest.json`

## Notes

- The add-in targets Outlook `Mailbox` requirement set `1.8`.
- Formatting is applied as HTML with `9pt` italic text to mimic the VBA output.
- New Outlook doesn't expose the Word editor object model used by VBA, so the add-in uses Office.js insertion APIs instead.
- The add-in now checks both true Outlook attachments and New Outlook shared file links. If a file was added with `Upload and share`, the add-in tries to read the visible filename from the link in the message body.
