using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToGoogleSync
{
    public partial class ThisAddIn
    {
        private CalendarService _calendarService;
        private Task _initTask;
        private readonly SemaphoreSlim _syncLock = new SemaphoreSlim(1, 1);

        // Keep strong references to Outlook.Items so .NET GC does not disconnect COM events
        private readonly List<Outlook.Items> _calendarItems = new List<Outlook.Items>();
        private readonly ConcurrentDictionary<string, string> _lastSyncedState = new ConcurrentDictionary<string, string>();

        private readonly string _appFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookToGoogleSync");

        // "primary" targets the main Google Calendar. Replace with a specific Calendar ID if needed.
        private const string TargetCalendarId = "primary";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Directory.CreateDirectory(_appFolder);
            Log("=== Starting OutlookToGoogleSync Add-in ===");

            _initTask = Task.Run(async () =>
            {
                try
                {
                    await InitializeGoogleCalendarAsync();
                }
                catch (Exception ex)
                {
                    Log($"Google API initialization failed: {ex}");
                }
            });

            Application.NewMailEx += Application_NewMailEx;
            HookAllCalendarFolders();
        }

        private async Task InitializeGoogleCalendarAsync()
        {
            string credPath = Path.Combine(_appFolder, "credentials.json");
            if (!File.Exists(credPath))
            {
                Log($"WARNING: credentials.json not found at: {credPath}");
                return;
            }

            UserCredential credential;
            using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
            {
                string tokenPath = Path.Combine(_appFolder, "token.json");
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] { CalendarService.Scope.CalendarEvents },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenPath, true));
            }

            _calendarService = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Outlook to Samsung Calendar Bridge"
            });

            Log("Google Calendar API connected successfully.");
        }

        private void HookAllCalendarFolders()
        {
            Outlook.Stores stores = null;
            try
            {
                stores = Application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    Outlook.MAPIFolder calFolder = null;
                    try
                    {
                        calFolder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                        if (calFolder != null)
                        {
                            Outlook.Items items = calFolder.Items;
                            items.ItemAdd += CalendarItems_ItemAdd;
                            items.ItemChange += CalendarItems_ItemChange;
                            _calendarItems.Add(items);
                            Log($"Monitoring calendar folder: '{store.DisplayName}' -> '{calFolder.Name}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Skipped calendar hook for store '{store.DisplayName}': {ex.Message}");
                    }
                    finally
                    {
                        if (calFolder != null) Marshal.ReleaseComObject(calFolder);
                        Marshal.ReleaseComObject(store);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Failed to enumerate Outlook stores: {ex.Message}");
            }
            finally
            {
                if (stores != null) Marshal.ReleaseComObject(stores);
            }
        }

        private void CalendarItems_ItemAdd(object item)
        {
            HandleCalendarAppointment(item, "Calendar ItemAdd");
        }

        private void CalendarItems_ItemChange(object item)
        {
            HandleCalendarAppointment(item, "Calendar ItemChange");
        }

        private void HandleCalendarAppointment(object item, string source)
        {
            if (item is Outlook.AppointmentItem appt)
            {
                try
                {
                    bool isCanceled = appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled ||
                                      appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;

                    var dto = BuildDtoFromAppointment(appt, appt.Organizer, isCanceled);
                    QueueSyncIfChanged(dto, source);
                }
                catch (Exception ex)
                {
                    Log($"Error handling {source}: {ex.Message}");
                }
            }
        }

        private void Application_NewMailEx(string entryIDCollection)
        {
            string[] entryIds = entryIDCollection.Split(',');
            Outlook.NameSpace session = Application.Session;

            foreach (string entryId in entryIds)
            {
                object item = null;
                try
                {
                    item = session.GetItemFromID(entryId.Trim());

                    if (item is Outlook.MeetingItem meetingItem)
                    {
                        ProcessMeetingItem(meetingItem);
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error processing item {entryId}: {ex.Message}");
                }
                finally
                {
                    if (item != null) Marshal.ReleaseComObject(item);
                }
            }
        }

        private void ProcessMeetingItem(Outlook.MeetingItem meetingItem)
        {
            string msgClass = meetingItem.MessageClass ?? "";
            Log($"Received MeetingItem: Class={msgClass}, Subject='{meetingItem.Subject}'");

            Outlook.AppointmentItem appt = null;
            try
            {
                bool isCancel = msgClass.IndexOf("Canceled", StringComparison.OrdinalIgnoreCase) >= 0;
                appt = meetingItem.GetAssociatedAppointment(!isCancel);
                if (appt == null)
                {
                    Log("Associated AppointmentItem not found.");
                    return;
                }

                string organizer = meetingItem.SenderName ?? meetingItem.SenderEmailAddress ?? appt.Organizer ?? "";
                var dto = BuildDtoFromAppointment(appt, organizer, isCancel);
                QueueSyncIfChanged(dto, "NewMailEx");
            }
            finally
            {
                if (appt != null) Marshal.ReleaseComObject(appt);
            }
        }

        private MeetingDto BuildDtoFromAppointment(Outlook.AppointmentItem appt, string fallbackOrganizer, bool isCanceled)
        {
            return new MeetingDto
            {
                GlobalId = appt.GlobalAppointmentID,
                Subject = string.IsNullOrWhiteSpace(appt.Subject) ? "(No Subject)" : appt.Subject,
                Location = appt.Location ?? "",
                Body = appt.Body ?? "",
                StartUtc = appt.StartUTC,
                EndUtc = appt.EndUTC,
                IsAllDay = appt.AllDayEvent,
                Organizer = string.IsNullOrWhiteSpace(fallbackOrganizer) ? (appt.Organizer ?? "") : fallbackOrganizer,
                IsCanceled = isCanceled
            };
        }

        private void QueueSyncIfChanged(MeetingDto dto, string source)
        {
            if (string.IsNullOrEmpty(dto.GlobalId))
                return;

            string fingerprint = $"{dto.Subject}|{dto.StartUtc:O}|{dto.EndUtc:O}|{dto.Location}|{dto.IsAllDay}|{dto.IsCanceled}|{dto.Body.GetHashCode()}";
            if (_lastSyncedState.TryGetValue(dto.GlobalId, out string previousFingerprint) && previousFingerprint == fingerprint)
            {
                return;
            }

            _lastSyncedState[dto.GlobalId] = fingerprint;
            Log($"Queuing sync ({source}): '{dto.Subject}' [{dto.StartUtc:yyyy-MM-dd HH:mm} UTC]");
            Task.Run(async () => await SyncGoogleEventAsync(dto));
        }

        private async Task SyncGoogleEventAsync(MeetingDto dto)
        {
            await _syncLock.WaitAsync();
            try
            {
                if (_initTask != null)
                {
                    await _initTask;
                }

                if (_calendarService == null)
                {
                    Log($"Skipping sync for '{dto.Subject}': Google Calendar service is not initialized.");
                    return;
                }

                var listReq = _calendarService.Events.List(TargetCalendarId);
                listReq.PrivateExtendedProperty = new Google.Apis.Util.Repeatable<string>(
                    new[] { $"outlookGlobalId={dto.GlobalId}" });

                var existingEvents = await listReq.ExecuteAsync();
                bool exists = existingEvents.Items != null && existingEvents.Items.Count > 0;

                if (dto.IsCanceled)
                {
                    if (exists)
                    {
                        string eventId = existingEvents.Items[0].Id;
                        await _calendarService.Events.Delete(TargetCalendarId, eventId).ExecuteAsync();
                        Log($"Event deleted (canceled): {dto.Subject}");
                    }
                    return;
                }

                var googleEvent = new Event
                {
                    Summary = dto.Subject,
                    Location = dto.Location,
                    Description = string.IsNullOrWhiteSpace(dto.Organizer)
                        ? dto.Body
                        : $"Organizer: {dto.Organizer}\n\n{dto.Body}",
                    ExtendedProperties = new Event.ExtendedPropertiesData
                    {
                        Private__ = new Dictionary<string, string>
                        {
                            { "outlookGlobalId", dto.GlobalId }
                        }
                    }
                };

                if (dto.IsAllDay)
                {
                    googleEvent.Start = new EventDateTime { Date = dto.StartUtc.ToString("yyyy-MM-dd") };
                    googleEvent.End = new EventDateTime { Date = dto.EndUtc.ToString("yyyy-MM-dd") };
                }
                else
                {
                    googleEvent.Start = new EventDateTime { DateTimeDateTimeOffset = new DateTimeOffset(DateTime.SpecifyKind(dto.StartUtc, DateTimeKind.Utc)) };
                    googleEvent.End = new EventDateTime { DateTimeDateTimeOffset = new DateTimeOffset(DateTime.SpecifyKind(dto.EndUtc, DateTimeKind.Utc)) };
                }

                if (exists)
                {
                    string eventId = existingEvents.Items[0].Id;
                    await _calendarService.Events.Update(googleEvent, TargetCalendarId, eventId).ExecuteAsync();
                    Log($"Event updated in Google Calendar: {dto.Subject}");
                }
                else
                {
                    await _calendarService.Events.Insert(googleEvent, TargetCalendarId).ExecuteAsync();
                    Log($"Event created in Google Calendar: {dto.Subject}");
                }
            }
            catch (Exception ex)
            {
                // Allow retry on next ItemChange if API call failed
                _lastSyncedState.TryRemove(dto.GlobalId, out _);
                Log($"Google Calendar API sync error for '{dto.Subject}': {ex.Message}");
            }
            finally
            {
                _syncLock.Release();
            }
        }

        private void Log(string message)
        {
            try
            {
                string logFile = Path.Combine(_appFolder, "sync.log");
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.NewMailEx -= Application_NewMailEx;
            foreach (var items in _calendarItems)
            {
                try
                {
                    items.ItemAdd -= CalendarItems_ItemAdd;
                    items.ItemChange -= CalendarItems_ItemChange;
                    Marshal.ReleaseComObject(items);
                }
                catch { }
            }
            _calendarItems.Clear();
            _syncLock.Dispose();
            Log("=== Stopping OutlookToGoogleSync Add-in ===");
        }

        private class MeetingDto
        {
            public string GlobalId { get; set; }
            public string Subject { get; set; }
            public string Location { get; set; }
            public string Body { get; set; }
            public DateTime StartUtc { get; set; }
            public DateTime EndUtc { get; set; }
            public bool IsAllDay { get; set; }
            public string Organizer { get; set; }
            public bool IsCanceled { get; set; }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}