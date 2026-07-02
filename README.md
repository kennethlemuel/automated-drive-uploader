<p align="center">
  <img src="assets/ESLogo.png" alt="Efficient Audit Uploader logo" width="160">
</p>

# Efficient Audit Uploader

Desktop app that uploads audit screenshots to Google Drive, writes their links into an `.xlsx` workbook, and can publish the finished workbook to Google Sheets.

## Download

Open the repository's **Releases** page and download the ZIP for your computer:

- `macOS-arm64` for Apple Silicon Macs (M1 or newer)
- `macOS-x64` for Intel Macs
- `Windows-x64` for 64-bit Windows

Extract the ZIP, then open **Efficient Audit Uploader**.

## First time opening the app

Your computer may show a warning because the app has not yet been registered with Apple or Microsoft. Only continue if you downloaded the ZIP from this repository's **Releases** page.

### Mac

1. Try to open **Efficient Audit Uploader**.
2. If your Mac blocks it, open **System Settings**.
3. Go to **Privacy & Security**.
4. Scroll down and click **Open Anyway**.
5. Enter your Mac password if asked.

You should only need to do this once for each downloaded version.

### Windows

1. Open **Efficient Audit Uploader.exe**.
2. If Windows shows a warning, click **More info**.
3. Click **Run anyway**.

If **Run anyway** is not available, ask your IT team for help. Your company may block apps that have not been approved.

## Google setup

1. Create or select a Google Cloud project.
2. Enable the Google Drive API.
3. Configure the OAuth consent screen and add coworkers as test users if the app is still in testing.
4. Create an OAuth client with application type **Desktop app**.
5. Download its JSON file. Each user selects this file under **Credentials** on first launch.

Credentials and OAuth tokens stay on the user's computer and are never bundled into a release.

## Use

1. Choose the credentials JSON.
2. Choose the audit workbook.
3. Choose the screenshots folder.
4. Set the Drive folder if needed.
5. Select the desired sharing and Sheets options, then click upload.

The original workbook is unchanged. The app writes a new `-with-links.xlsx` file unless the output is renamed.

## For developers

Coworkers who downloaded the app can skip this section. These steps are only for people changing or rebuilding the app.

Python 3.11 is used for release builds.

```sh
python -m pip install -r requirements-dev.txt
python test_excel_drive_links.py
python desktop_app.py
```

Build the native package for the current operating system:

```sh
python build.py
```

The result is written to `dist/`.

## Publish a release

Push a version tag to GitHub:

```sh
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions builds the three platform packages and attaches them to the new GitHub Release. The workflow can also be run manually to create downloadable Actions artifacts without publishing a release.

These opening warnings are acceptable for internal testing. They can be removed later by registering the app with Apple and Microsoft.
