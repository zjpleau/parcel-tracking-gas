# Parcel Tracking Sync (Google Apps Script)

A modern, automated Google Apps Script project that scans your Gmail for shipping confirmation emails and automatically syncs tracking numbers to the **Parcel API**.

## 🚀 Features

- **Automated Scanning**: Scans Gmail threads for UPS, USPS, FedEx, and OnTrac tracking numbers.
- **Smart Parsing**: Intelligent extraction of merchant names and package descriptions.
- **Quota Management**: Built-in daily rate limiting to respect API tiers.
- **Daily Summaries**: Optional daily email reports of all packages processed.
- **Error Handling**: Graceful error management and notifications.

## 🛠 Project Structure

- `Code.js`: Main logic for scanning, parsing, and API synchronization.
- `appsscript.json`: Manifest file for the Google Apps Script project.
- `.clasp.json`: Configuration for the `clasp` CLI to sync with Google.

## 💻 Local Development Workflow

This project is set up for local development using [clasp](https://github.com/google/clasp).

### Pushing Changes to Google
To push your local changes to the Google Apps Script project:
```bash
clasp push
```

### Pulling Changes from Google
If you made changes directly in the Google Script editor:
```bash
clasp pull
```

### Version Control
- Use **Git** for logic changes and refactoring.
- Create **Pull Requests** for significant features.
- Use **GitHub Issues** to track bugs or planned improvements.

## ⚙️ Configuration

Ensure the following Script Properties are set in your Google Project:
- `PARCEL_API_KEY`: Your Parcel API key.
- `PARCEL_API_URL`: (Optional) Custom API endpoint.

---
*Created and maintained with Antigravity.*
