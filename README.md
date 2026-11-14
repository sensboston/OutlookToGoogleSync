# OutlookToGoogleSync

[![License: MIT](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=flat-square&logo=windows)
![Outlook: Classic Desktop](https://img.shields.io/badge/Outlook-Classic%20Desktop-0078D4?style=flat-square)
![.NET Framework: 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat-square&logo=dotnet)
![GitHub last commit](https://img.shields.io/github/last-commit/sensboston/OutlookToGoogleSync/master?style=flat-square)

A lightweight, 100% private Windows VSTO Add-in for **Classic Microsoft Outlook** that automatically synchronizes calendar events and meeting invitations directly to **Google Calendar** (and consequently to **Samsung Calendar** and Android home-screen widgets) without third-party cloud relays.

---

## The Problem

If you use desktop Outlook with **IMAP accounts** (personal domains, ISP email, standard hosting) alongside an **Android / Samsung Galaxy** phone, you face a well-known ecosystem gap:

1. **IMAP does not sync calendars:** Meeting invitations (`.ics`) arriving via IMAP are saved by desktop Outlook into a local *"This computer only"* calendar folder. They never reach the cloud.
2. **Broken mobile widget sync:** Even when using Exchange/Microsoft 365 alongside Outlook Mobile, Samsung Calendar relies on Android's `CalendarProvider`. Background sync between Outlook Mobile and Samsung Calendar frequently stalls due to One UI battery optimizations or app updates—leaving your Samsung Calendar widget empty until right before a meeting starts.
3. **Privacy & subscription bloat:** Existing commercial bridges require recurring fees or route your private work schedule through third-party servers.

## The Solution

**OutlookToGoogleSync** runs locally inside desktop Outlook and acts as a real-time, one-way bridge to your Google Calendar using your own Google Cloud OAuth credentials:

* **Zero third-party servers:** Data flows strictly from your PC directly to Google's official Calendar API.
* **Catches everything automatically:**
  * **Incoming invites & cancellations (`Application.NewMailEx`):** Instantly parses incoming `MeetingItem` requests and cancellations across all POP/IMAP/Exchange accounts before you even accept them.
  * **Manual events & updates (`Items.ItemAdd` / `Items.ItemChange`):** Hooks into all Outlook store calendar folders so events you create, edit, or drag-and-drop manually on your desktop are synced immediately.
* **Duplicate & spam prevention:** Maps Outlook's `GlobalAppointmentID` to Google Calendar `PrivateExtendedProperty` (`outlookGlobalId`), guarded by a `SemaphoreSlim` queue and an in-memory state fingerprint cache (`ConcurrentDictionary`) to prevent duplicate events or redundant API calls.
* **Native Samsung Calendar integration:** Once an event lands in Google Calendar, Android's native Google sync adapter populates `CalendarProvider`, making it appear automatically in the Samsung Calendar app and widgets.

---

## Prerequisites

* Windows 10 / 11 with **Classic Microsoft Outlook** (Office 2016, 2019, 2021, LTSC, or Microsoft 365 desktop). *Note: Does not support the web-based "New Outlook".*
* **Visual Studio 2022** with the **Office/SharePoint development** workload installed.
* A Google account (the same one connected to your Android / Samsung phone).

---

## File Locations & Architecture

All runtime configuration, OAuth tokens, and logs are stored locally in your Windows user profile directory (never inside the repository or compiled binaries):

```text
%AppData%\OutlookToGoogleSync\
├── credentials.json   # Your Google Cloud OAuth 2.0 Desktop Client secret (you place this here)
├── token.json\        # Auto-generated folder storing your persistent OAuth refresh token
└── sync.log           # Execution, folder hook, and Google Calendar API sync logs
