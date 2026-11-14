# OutlookToGoogleSync

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

## Step-by-Step Setup Guide

### Step 1: Configure Google Cloud OAuth Credentials

Because this is a self-hosted personal bridge, you create your own free Google Cloud OAuth client so your data never touches someone else's app.

1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and create a new project (e.g., `OutlookCalendarSync`).
2. Navigate to **APIs & Services → Library**, search for **Google Calendar API**, and click **Enable**.
3. Open **Google Auth Platform** (or *OAuth consent screen*):
   * **Branding:** Fill in *App name*, *User support email*, and *Developer contact email*. Under *App domain*, enter your website or placeholder HTTPS URLs (e.g., `https://example.com`) and add the domain to *Authorized domains*. **Do not upload an App logo** (uploading a logo triggers mandatory Google app review). Click **Save**.
   * **Audience:** Select **External**, then click **Publish App → Confirm** to switch the status from *Testing* to **In production**. *(Crucial: Apps left in "Testing" expire OAuth refresh tokens every 7 days; "In production" keeps your token valid indefinitely without requiring Google verification for personal use under 100 users).*
4. Go to **Clients** (or *Credentials*) → **Create Client → Desktop app**, give it a name, and click **Create**.
5. Download the generated JSON file, rename it to `credentials.json`, and place it in:
   ```text
   %AppData%\OutlookToGoogleSync\credentials.json
