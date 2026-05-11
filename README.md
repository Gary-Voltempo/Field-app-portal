# Voltempo Field Application Web Portal

Static web portal for reviewing, editing, submitting, and approving Voltempo
commissioning reports.

## Run locally

Serve the repo root so the portal can load shared logo assets:

```sh
python3 -m http.server 4173
```

Open:

```text
http://localhost:4173/desktop_web/
```

## Microsoft app registration

The portal uses the web app registration client ID:

```text
28753d72-1ee9-450b-9946-43247f9b0a24
```

Add this as a Single-page application redirect URI in Microsoft Entra:

```text
http://localhost:4173/desktop_web/
```

When deployed to Azure Static Web Apps, also add the production URL as a
Single-page application redirect URI, for example:

```text
https://<your-static-web-app-name>.azurestaticapps.net/
```

The required delegated Microsoft Graph scopes are:

```text
User.Read
Files.ReadWrite
Mail.Send
```

The app reads and writes this OneDrive file:

```text
voltempo_commissioning_reports.json
```

## Azure Static Web Apps deployment

Recommended Azure setup:

```text
App location: /desktop_web
Output location: leave blank
API location: leave blank
Build preset: Custom / HTML
```

After Azure creates the Static Web App:

1. Copy the generated Azure URL.
2. Add it to the Microsoft Entra app registration as a SPA redirect URI.
3. Open the Azure URL and sign in as `commissioning@voltempo.com`.
4. Consent to these delegated Graph permissions if prompted:
   `User.Read`, `Files.ReadWrite`, `Mail.Send`.
5. Submit a test report and confirm Gary receives the notification email.

The app includes `staticwebapp.config.json` so Azure serves the static portal
cleanly and falls back to `index.html` if needed.

## Next slice

Move the report data source from the commissioning account OneDrive into a
SharePoint document library once the workflow is stable.
