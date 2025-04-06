# ExchangeMailFlowMonitor.ps1

🔄 **Real-time Exchange Server 2016 Mail Flow & Queue Monitor**  
PowerShell script for live monitoring of inbound/outbound email traffic and mail queues on Exchange Server 2016.

## 📚 Table of Contents
- [Overview](#-overview)
- [Requirements](#️-requirements)
- [Script Configuration](#-script-configuration)
- [How to Use](#-how-to-use)
- [Output Example](#-output-example)
- [Notes & Considerations](#-notes--considerations)
- [Author](#-author)
- [License](#-license)

---

## 📋 Overview

This script provides a **real-time snapshot** of:

- The last 10 inbound messages received
- The last 10 outbound messages sent
- Current queue status (active, retrying, etc.)
- Messages currently waiting in queues

It refreshes the output every few seconds (default: 10), giving system administrators a live dashboard for troubleshooting and visibility into message flow.

---

## ⚙️ Requirements

- **Exchange Server 2016** (may also work on 2013/2019 with minor adjustments)
- **PowerShell with appropriate permissions**
- Must be run from a machine with the **Exchange Management Tools** installed (or directly on the Exchange server)

---

## 🧾 Script Configuration

You can configure the following parameters at the top of the script:

```powershell
$server = "EXCHANGE01"         # Your Exchange server name
$refreshInterval = 10          # Seconds between refreshes
$lookBackMinutes = 5           # How far back to look in logs (minutes)
```

---

## 🚀 How to Use

1. Open PowerShell **as Administrator**.
2. Run the script:
    ```powershell
    .\ExchangeMailFlowMonitor.ps1
    ```
3. Watch live updates of mail flow and queues.

> 🔄 Press `CTRL+C` at any time to stop the monitoring loop.

---

## 📊 Output Example

```text
===== Exchange Mail Traffic + Queue Monitor on EXCHANGE01 (Updated: 2025-04-06 10:22:31) =====

--- LAST 10 INBOUND MAILS ---
Timestamp            Sender               Recipients                 MessageSubject
---------            ------               ----------                 ---------------
2025-04-06 10:21     user1@domain.com     helpdesk@domain.com       System Alert
...

--- LAST 10 OUTBOUND MAILS ---
Timestamp            Sender               Recipients                 MessageSubject
---------            ------               ----------                 ---------------
2025-04-06 10:20     alerts@domain.com    admin@domain.com          Notification
...

--- CURRENT MAIL QUEUES ---
Identity             Status   MsgCount    DeliveryType     NextHopDomain
--------             ------   --------    ------------     --------------
EXCHANGE01\Queue1    Active     4         SmtpRelay        remote.domain.com

--- MESSAGES IN QUEUES ---
Queue                FromAddress          Subject             Status       DateReceived
-----                ------------         -------             ------       -------------
EXCHANGE01\Queue1    alerts@domain.com    System Warning      Ready        2025-04-06 10:19
```

---

## ⚠️ Notes & Considerations

- The script uses `Get-MessageTrackingLog`, `Get-Queue`, and `Get-Message` — make sure your user account has the necessary **Exchange permissions**.
- Output may vary based on your organization's message tracking and retention settings.
- If you're running this in a multi-server environment, consider adapting the script to iterate over multiple servers.

---

## 🧑‍💻 Author

Created by aviado1  
[GitHub Profile](https://github.com/aviado1)

