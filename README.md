# SIFY Migration Tool — Online Module 1

Online Flask deployment for **Module 1: Basic Configuration only**.

- Branch lookup from Excel
- Three live-link scenarios
- Review generated FortiGate configuration
- Download generated configuration
- No Netmiko / live FortiGate / NMS execution in this online version
- No device credentials stored in source

## Render
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Set `BASIC_USER`, `BASIC_PASSWORD`, `FLASK_SECRET`, and `EXCEL_FILE` as environment variables.

**Security:** keep this repository private before adding the migration workbook. The workbook contains internal branch/IP/network information.
