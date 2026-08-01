# OutlookQuickMove Installation and User Guide

English | [中文](USER_GUIDE.zh-CN.md)

OutlookQuickMove is a VSTO add-in for classic Outlook for Windows. It helps you move selected mail items to a searched folder, jump to a folder by keyword, or jump to the folder that contains a selected mail item.

## Requirements

- Classic Outlook for Windows.
- New Outlook, Outlook on the web, and Outlook for Mac are not supported.
- Close Outlook before installing the add-in.

## Install

1. Open the GitHub Releases page and download the latest release archive, for example `OutlookQuickMove-1.0.0.9.7z`.
2. Extract the archive to a stable local folder. Do not run the installer from inside the archive.
3. Make sure Outlook is closed.
4. Double-click `setup.exe` from the extracted folder.
5. Follow the installation wizard.
6. Open Outlook after installation completes.

![OutlookQuickMove release entry on GitHub Releases](assets/user-guide/01-release-download.png)

## Open Quick Move

After installation and restarting Outlook, a `Schnell verschieben` tab appears in the Outlook ribbon.

![Quick Move ribbon tab in Outlook](assets/user-guide/02-quick-move-tab.png)

The tab contains these buttons:

- `Schnell verschieben`
- `Go to Folder`
- `Go to Mail Folder`
- `Undo Quick Move...`
- `Settings`

## Quick Move

`Quick Move` moves the currently selected mail items to a target folder.

1. Select one or more mail items in Outlook.
2. Press `Ctrl+M`, Thunderbird-style `Shift+M`, or click `Schnell verschieben`.
3. Type keywords from the target account, folder name, or path.
4. Use `Up` / `Down` to choose a folder.
5. Press `Enter` to move the selected mail items.

The picker searches across all enabled Outlook accounts and data files. It groups matching folders
under compact account headers and displays the parent path on the right. Your configured Outlook
accounts are listed in Outlook's account order before additional/shared mailboxes. With an empty
search box, the picker shows your recently and frequently used targets. Use `Ctrl+Left` /
`Ctrl+Right` to switch between `Verschieben`, `Kopieren`, and `Zum Ordner` without leaving the
keyboard.

The folder index is cached in memory for 10 minutes, kept on disk for up to 30 days, and preloaded
when the add-in starts. Press `Aktualisieren` after adding or renaming folders if you need to rebuild
it immediately.

Select `Als gelesen verschieben` if the selected mail should be marked as read before it is moved.

![Quick Move dialog and filtered folder results](assets/user-guide/03-quick-move-dialog.png)

## Go to Folder

`Go to Folder` jumps to a folder without moving any mail.

1. Press `Ctrl+G`, Thunderbird-style `Shift+G`, or click `Go to Folder`.
2. Type folder keywords.
3. Select the target folder.
4. Press `Enter`. Outlook switches to that folder.

It uses the same folder scope as `Quick Move` and reuses the frequent-folder ordering.

## Go to Mail Folder

`Go to Mail Folder` jumps to the folder that contains the currently selected mail item.

This is useful from Outlook search results. After finding a message, select it and click `Go to Mail Folder` to return to the original folder that contains it.

If multiple items are selected, the add-in asks for confirmation and uses the first selected mail item as the target.

## Undo Quick Move

`Undo Quick Move...` reverses moves made by `Quick Move`.

1. Click `Undo Quick Move...`.
2. Choose the move records you want to undo.
3. Click `Undo Selected`.
4. The add-in tries to move the mail items back to their original folders and restore their previous read or unread state.

![Undo Quick Move dialog](assets/user-guide/04-undo-quick-move.png)

## Settings

`Settings` controls the searched data files, frequent-folder count, and undo-history count.

### Data Files

Use `Data Files` to choose which Outlook data files are searched for target folders.

If you only want to search one mailbox or PST file, select only that data file. This keeps the candidate list smaller and can reduce the cost of building the folder list the first time.

![Data Files tab in Settings](assets/user-guide/05-settings-data-files.png)

### Frequent Folders

`Frequent Folders` controls how many frequently used target folders appear at the top of the candidate list.

Frequent folders are ordered automatically based on actual `Quick Move` usage. Set the value to `0` to turn off the frequent-folder display.

![Frequent Folders tab in Settings](assets/user-guide/06-settings-frequent-folders.png)

### Undo History

`Undo History` controls how many recent move records are kept for undo.

Set the value to `0` to stop recording new undo history. Existing history remains until it is cleared manually.

![Undo History tab in Settings](assets/user-guide/07-settings-undo-history.png)

## Keyboard Shortcuts

- `Ctrl+M` or `Shift+M`: open Quick Move with `Verschieben` selected.
- `Ctrl+G` or `Shift+G`: open the folder picker in navigation mode.
- Type, use `Up` / `Down`, then press `Enter`.
- `Ctrl+Left` / `Ctrl+Right`: switch the action in the unified Quick Move dialog.
- `Esc`: close the picker without changing anything.
- Clicking outside the compact picker closes it without changing anything.

The shortcuts are scoped to the classic Outlook UI thread and are not system-wide. Classic Outlook
normally uses `Ctrl+M` for Send/Receive; this fork intentionally reserves it for Quick Move. Use
`F9` for Send/Receive. The `Shift` aliases are active only in the Outlook Explorer while no text
field has an active caret, so typing uppercase `M` or `G` in searches and editors is unaffected.

## Screenshot Privacy Check

Before syncing these screenshots to the public repository, confirm that they do not include:

- Real names, email addresses, work accounts, or profile photos.
- Mail subjects, sender names, recipient names, or message previews.
- Private mailbox, PST, OST, or shared mailbox names.
- Private folder names, customer names, project names, or ticket numbers.
- Windows usernames, local paths, or network paths.
- Any detail that can identify a person, workplace, customer, or private workflow.

Use test mailboxes and test folders where possible. If a real Outlook profile must be shown, blur or crop every sensitive area before committing the screenshot.
