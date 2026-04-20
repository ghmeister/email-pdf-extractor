# Email PDF Extractor

A Dockerized Python service that scans an IMAP inbox, extracts PDF attachments, applies configurable rules, and uploads matching PDF files to OneDrive.

## Features

- IMAP polling for new emails
- configurable rule engine for sender, subject, filename, email body, and PDF text content
- OneDrive upload using Microsoft Graph
- simple web UI for rule management and log viewing
- Docker-ready for Portainer deployment

## Quickstart

1. Copy the project to your Git folder.
2. Create `data` folder next to `docker-compose.yml`.
3. Set environment variables in Portainer or a `.env` file:
   - `IMAP_HOST`
   - `IMAP_PORT` (default `993`)
   - `IMAP_USER`
   - `IMAP_PASS`
   - `IMAP_FOLDER` (default `INBOX`)
   - `POLL_INTERVAL` (default `120` seconds)
   - `ONEDRIVE_TENANT_ID`
   - `ONEDRIVE_CLIENT_ID`
   - `ONEDRIVE_CLIENT_SECRET`
   - `ONEDRIVE_REFRESH_TOKEN`
   - `ONEDRIVE_FOLDER_PATH` (default `/Documents/Invoices`)
   - `SECRET_KEY`

4. Start the container with Portainer or `docker compose up -d`.
5. Open the UI at `http://<host>:5001` to manage rules and view logs.

## Rules

The web UI lets you add rules to match PDFs based on:

- sender contains
- subject contains
- attachment filename contains
- email body contains
- PDF text contains

## Notes

- The app stores rules and logs in SQLite under `data/app.db`.
- OneDrive upload requires a valid OAuth refresh token and Azure app credentials.
- The initial default rule matches invoice-like messages.
