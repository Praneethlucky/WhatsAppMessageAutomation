# WhatsApp Bulk Sender (Windows)

## Features
- Select an Excel workbook with `Contacts` and `Messages` sheets.
- `Contacts` columns: `Phone`, `Name`, `OptIn`.
- `Messages` column: `Message`.
- Supports `{name}` placeholder in the message.
- Only processes rows where `OptIn` is `Yes`, `True`, or `1`.
- Start and stop controls, progress display, timestamped CSV log.
- No explicit inter-message sleep/interval. PyWhatKit's `wait_time=20` remains as a browser-loading wait.

## Excel format

**Contacts** sheet:

| Phone | Name | OptIn |
|---|---|---|
| +919876543210 | Example | No |

**Messages** sheet:

| Message |
|---|
| Hello {name}, this is an example. |

Only change OptIn to Yes for people who explicitly agreed to receive WhatsApp messages.

## Run from Python
```bash
python -m pip install -r requirements.txt
python main.py
```

## Build Windows EXE
Run `build_exe.bat` on a Windows computer. The executable will be:
`dist\WhatsAppBulkSender.exe`

Build on Windows for Windows. The target computer does not need Python, but it does need:
- A compatible Windows version
- Internet access
- A supported browser installed
- WhatsApp Web logged in

## Logs
CSV logs are created in a `logs` folder beside the selected workbook. Status `SUBMITTED` means the PyWhatKit call returned; it does not prove delivery or reading.

## Important
PyWhatKit automates WhatsApp Web and is not an official WhatsApp Business API. Browser automation can fail or be blocked, and sending unsolicited bulk messages may lead to account restrictions. Use only for opted-in recipients, comply with WhatsApp policies, and honor opt-outs. Test with a small group first.
