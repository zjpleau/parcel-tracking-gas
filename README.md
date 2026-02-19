# Parcel Tracking Sync

A Google Apps Script utility that scans Gmail for shipping confirmation emails and syncs tracking numbers to the Parcel API.

## Features

- **Automated Scanning**: Identifies tracking numbers from UPS, USPS, FedEx, and OnTrac.
- **Data Extraction**: Extracts merchant names and shipment descriptions.
- **Rate Limiting**: Includes daily rate limits for API management.
- **Reporting**: Generates daily email summaries of processed packages.

## Structure

- `Code.js`: Main logic for email scanning and API synchronization.
- `appsscript.json`: Google Apps Script manifest.
- `.clasp.json`: CLI configuration for synchronizing with the Google script environment.

## Development Workflow

This project uses [clasp](https://github.com/google/clasp) for local development and deployment.

### Deploying Changes
Push local changes to Google Apps Script:
```bash
clasp push
```

### Retrieving Changes
Pull changes from the Google Script editor:
```bash
clasp pull
```

### Version Control
- Use Git for logic changes and project history.
- Use GitHub Issues for bug tracking and feature planning.

## Configuration

This utility integrates with [Parcel App](https://web.parcelapp.net/). You must have a registered account to obtain an API key.

Required Script Properties:
- `PARCEL_API_KEY`: Your API key from the Parcel App settings.
- `PARCEL_API_URL`: (Optional) Target API endpoint.
