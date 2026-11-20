using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using TimeZoneConverter;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToGoogleSync
{
    public partial class ThisAddIn
    {
        private CalendarService _calendarService;
        private Task _initTask;
        private readonly SemaphoreSlim _syncLock = new SemaphoreSlim(1, 1);

        // Keep strong references to Outlook.Items and Outlook.Folder so .NET GC does not disconnect COM events
        private readonly List<Outlook.Items> _calendarItems = new List<Outlook.Items>();
        private readonly List<Outlook.Folder> _calendarFolders = new List<Outlook.Folder>();
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

                            // Deleting an appointment (Delete, Shift+Delete, declining a meeting) moves it out of the
                            // calendar; Items.ItemRemove does not say which item left, BeforeItemMove does.
                            if (calFolder is Outlook.Folder folder)
                            {
                                folder.BeforeItemMove += CalendarFolder_BeforeItemMove;
                                _calendarFolders.Add(folder);
                                calFolder = null; // kept alive in _calendarFolders
                            }
                            Log($"Monitoring calendar folder: '{store.DisplayName}'");
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

        private void CalendarFolder_BeforeItemMove(object item, Outlook.MAPIFolder moveTo, ref bool cancel)
        {
            try
            {
                // A move to another calendar is picked up there by ItemAdd; anything else (Deleted Items,
                // a hard delete with moveTo == null) removes the event.
                if (moveTo != null && moveTo.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                    return;

                if (item is Outlook.AppointmentItem appt)
                {
                    string globalId = appt.GlobalAppointmentID;
                    string subject = appt.Subject ?? "(No Subject)";
                    if (string.IsNullOrEmpty(globalId))
                        return;

                    _lastSyncedState.TryRemove(globalId, out _);
                    Log($"Queuing delete (removed from calendar): '{subject}'");
                    Task.Run(async () => await DeleteGoogleEventAsync(globalId, subject));
                }
            }
            catch (Exception ex)
            {
                Log($"Error handling BeforeItemMove: {ex.Message}");
            }
        }

        private void HandleCalendarAppointment(object item, string source)
        {
            if (item is Outlook.AppointmentItem appt)
            {
                try
                {
                    appt = SeriesMaster(appt);
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

        // An occurrence or an exception stands for its series: the whole series is synced from its master.
        private static Outlook.AppointmentItem SeriesMaster(Outlook.AppointmentItem appt)
        {
            if (appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence ||
                appt.RecurrenceState == Outlook.OlRecurrenceState.olApptException)
            {
                if (appt.GetRecurrencePattern().Parent is Outlook.AppointmentItem master)
                    return master;
            }
            return appt;
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

                // A cancellation of one occurrence arrives as an exception of the series: sync the series.
                var master = SeriesMaster(appt);
                bool wholeCancel = isCancel && ReferenceEquals(master, appt);

                string organizer = meetingItem.SenderName ?? meetingItem.SenderEmailAddress ?? master.Organizer ?? "";
                var dto = BuildDtoFromAppointment(master, organizer, wholeCancel);
                QueueSyncIfChanged(dto, "NewMailEx");
            }
            finally
            {
                if (appt != null) Marshal.ReleaseComObject(appt);
            }
        }

        private MeetingDto BuildDtoFromAppointment(Outlook.AppointmentItem appt, string fallbackOrganizer, bool isCanceled)
        {
            var dto = new MeetingDto
            {
                GlobalId = appt.GlobalAppointmentID,
                Subject = string.IsNullOrWhiteSpace(appt.Subject) ? "(No Subject)" : appt.Subject,
                Location = appt.Location ?? "",
                Body = appt.Body ?? "",
                StartUtc = appt.StartUTC,
                EndUtc = appt.EndUTC,
                StartLocal = appt.Start,
                EndLocal = appt.End,
                IsAllDay = appt.AllDayEvent,
                Organizer = string.IsNullOrWhiteSpace(fallbackOrganizer) ? (appt.Organizer ?? "") : fallbackOrganizer,
                IsCanceled = isCanceled
            };

            if (appt.IsRecurring && !isCanceled)
            {
                try
                {
                    FillRecurrence(appt, dto);
                }
                catch (Exception ex)
                {
                    // Better a single event than none: the first occurrence is still synced.
                    Log($"Recurrence of '{dto.Subject}' not readable, syncing the first occurrence only: {ex.Message}");
                    dto.Recurrence = null;
                    dto.Modified.Clear();
                }
            }
            return dto;
        }

        // Translates the Outlook recurrence pattern into an RFC 5545 RRULE, deleted occurrences into EXDATE
        // and modified occurrences into instance overrides.
        private static void FillRecurrence(Outlook.AppointmentItem appt, MeetingDto dto)
        {
            Outlook.RecurrencePattern rp = appt.GetRecurrencePattern();
            dto.TimeZone = IanaTimeZone(appt);

            // The series starts at its pattern start; appt.Start of a master is the first occurrence.
            DateTime firstLocal = rp.PatternStartDate.Date + rp.StartTime.TimeOfDay;
            DateTime firstEndLocal = firstLocal.AddMinutes(rp.Duration);
            dto.StartLocal = dto.IsAllDay ? rp.PatternStartDate.Date : firstLocal;
            dto.EndLocal = dto.IsAllDay ? rp.PatternStartDate.Date.AddDays(Math.Max(1, rp.Duration / 1440)) : firstEndLocal;
            dto.StartUtc = dto.StartLocal.ToUniversalTime();
            dto.EndUtc = dto.EndLocal.ToUniversalTime();

            var rule = new StringBuilder("RRULE:");
            int interval = Math.Max(1, rp.Interval);
            switch (rp.RecurrenceType)
            {
                case Outlook.OlRecurrenceType.olRecursDaily:
                    rule.Append("FREQ=DAILY");
                    break;
                case Outlook.OlRecurrenceType.olRecursWeekly:
                    rule.Append("FREQ=WEEKLY;BYDAY=").Append(Days(rp.DayOfWeekMask));
                    break;
                case Outlook.OlRecurrenceType.olRecursMonthly:
                    rule.Append("FREQ=MONTHLY;BYMONTHDAY=").Append(rp.DayOfMonth);
                    break;
                case Outlook.OlRecurrenceType.olRecursMonthNth:
                    rule.Append("FREQ=MONTHLY;BYDAY=").Append(Days(rp.DayOfWeekMask)).Append(";BYSETPOS=").Append(Nth(rp.Instance));
                    break;
                case Outlook.OlRecurrenceType.olRecursYearly:
                    interval = Math.Max(1, rp.Interval / 12); // Outlook counts yearly intervals in months
                    rule.Append("FREQ=YEARLY;BYMONTH=").Append(rp.MonthOfYear).Append(";BYMONTHDAY=").Append(rp.DayOfMonth);
                    break;
                case Outlook.OlRecurrenceType.olRecursYearNth:
                    interval = Math.Max(1, rp.Interval / 12);
                    rule.Append("FREQ=YEARLY;BYMONTH=").Append(rp.MonthOfYear).Append(";BYDAY=").Append(Days(rp.DayOfWeekMask))
                        .Append(";BYSETPOS=").Append(Nth(rp.Instance));
                    break;
                default:
                    throw new NotSupportedException("recurrence type " + rp.RecurrenceType);
            }
            if (interval > 1) rule.Append(";INTERVAL=").Append(interval);
            if (!rp.NoEndDate) rule.Append(";COUNT=").Append(rp.Occurrences);

            var recurrence = new List<string> { rule.ToString() };

            foreach (Outlook.Exception ex in rp.Exceptions)
            {
                // OriginalDate is the date of the occurrence as the pattern placed it, in local time.
                DateTime originalLocal = ex.OriginalDate.Date + (dto.IsAllDay ? TimeSpan.Zero : rp.StartTime.TimeOfDay);
                string key = dto.IsAllDay
                    ? originalLocal.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    : originalLocal.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

                if (ex.Deleted)
                {
                    recurrence.Add(dto.IsAllDay ? "EXDATE;VALUE=DATE:" + key : "EXDATE:" + key);
                    continue;
                }

                Outlook.AppointmentItem changed = ex.AppointmentItem;
                dto.Modified.Add(new OccurrenceDto
                {
                    InstanceKey = key,
                    Subject = string.IsNullOrWhiteSpace(changed.Subject) ? dto.Subject : changed.Subject,
                    Location = changed.Location ?? "",
                    StartUtc = changed.StartUTC,
                    EndUtc = changed.EndUTC,
                    StartLocal = changed.Start,
                    EndLocal = changed.End
                });
            }
            dto.Recurrence = recurrence;
        }

        private static string Days(Outlook.OlDaysOfWeek mask)
        {
            var days = new List<string>();
            if ((mask & Outlook.OlDaysOfWeek.olMonday) != 0) days.Add("MO");
            if ((mask & Outlook.OlDaysOfWeek.olTuesday) != 0) days.Add("TU");
            if ((mask & Outlook.OlDaysOfWeek.olWednesday) != 0) days.Add("WE");
            if ((mask & Outlook.OlDaysOfWeek.olThursday) != 0) days.Add("TH");
            if ((mask & Outlook.OlDaysOfWeek.olFriday) != 0) days.Add("FR");
            if ((mask & Outlook.OlDaysOfWeek.olSaturday) != 0) days.Add("SA");
            if ((mask & Outlook.OlDaysOfWeek.olSunday) != 0) days.Add("SU");
            return string.Join(",", days);
        }

        // Outlook numbers the week of the month 1..5, where 5 means the last one.
        private static int Nth(int instance) => instance >= 5 ? -1 : instance;

        // Google needs an IANA zone to expand a series across daylight saving changes.
        private static string IanaTimeZone(Outlook.AppointmentItem appt)
        {
            string windowsId = null;
            try { windowsId = appt.StartTimeZone?.ID; } catch { }
            if (string.IsNullOrEmpty(windowsId)) windowsId = TimeZoneInfo.Local.Id;
            return TZConvert.TryWindowsToIana(windowsId, out string iana) ? iana : "UTC";
        }

        private void QueueSyncIfChanged(MeetingDto dto, string source)
        {
            if (string.IsNullOrEmpty(dto.GlobalId))
                return;

            string fingerprint = $"{dto.Subject}|{dto.StartUtc:O}|{dto.EndUtc:O}|{dto.Location}|{dto.IsAllDay}|{dto.IsCanceled}|{dto.Body.GetHashCode()}"
                + (dto.Recurrence == null ? "" : "|" + string.Join(";", dto.Recurrence))
                + string.Concat(dto.Modified.Select(m => $"|{m.InstanceKey}>{m.Subject}@{m.StartUtc:O}-{m.EndUtc:O}/{m.Location}"));
            if (_lastSyncedState.TryGetValue(dto.GlobalId, out string previousFingerprint) && previousFingerprint == fingerprint)
            {
                return;
            }

            _lastSyncedState[dto.GlobalId] = fingerprint;
            Log($"Queuing sync ({source}): '{dto.Subject}' [{dto.StartUtc:yyyy-MM-dd HH:mm} UTC]"
                + (dto.Recurrence == null ? "" : $" series {dto.Recurrence[0]}, {dto.Recurrence.Count - 1} deleted, {dto.Modified.Count} modified"));
            Task.Run(async () => await SyncGoogleEventAsync(dto));
        }

        private async Task<IList<Event>> FindGoogleEventsAsync(string globalId)
        {
            var listReq = _calendarService.Events.List(TargetCalendarId);
            listReq.PrivateExtendedProperty = new Google.Apis.Util.Repeatable<string>(
                new[] { $"outlookGlobalId={globalId}" });
            var found = await listReq.ExecuteAsync();
            return found.Items ?? new List<Event>();
        }

        private async Task<bool> ReadyAsync(string subject)
        {
            if (_initTask != null)
            {
                await _initTask;
            }
            if (_calendarService == null)
            {
                Log($"Skipping sync for '{subject}': Google Calendar service is not initialized.");
                return false;
            }
            return true;
        }

        private async Task SyncGoogleEventAsync(MeetingDto dto)
        {
            await _syncLock.WaitAsync();
            try
            {
                if (!await ReadyAsync(dto.Subject))
                    return;

                var existingEvents = await FindGoogleEventsAsync(dto.GlobalId);
                bool exists = existingEvents.Count > 0;

                if (dto.IsCanceled)
                {
                    foreach (var existing in existingEvents)
                    {
                        await _calendarService.Events.Delete(TargetCalendarId, existing.Id).ExecuteAsync();
                    }
                    if (exists) Log($"Event deleted (canceled): {dto.Subject}");
                    return;
                }

                var googleEvent = new Event
                {
                    Summary = dto.Subject,
                    Location = dto.Location,
                    Description = string.IsNullOrWhiteSpace(dto.Organizer)
                        ? dto.Body
                        : $"Organizer: {dto.Organizer}\n\n{dto.Body}",
                    Recurrence = dto.Recurrence,
                    ExtendedProperties = new Event.ExtendedPropertiesData
                    {
                        Private__ = new Dictionary<string, string>
                        {
                            { "outlookGlobalId", dto.GlobalId }
                        }
                    }
                };
                SetTimes(googleEvent, dto.IsAllDay, dto.StartLocal, dto.EndLocal, dto.StartUtc, dto.EndUtc, dto.TimeZone);

                Event saved;
                if (exists)
                {
                    saved = await _calendarService.Events.Update(googleEvent, TargetCalendarId, existingEvents[0].Id).ExecuteAsync();
                    Log($"Event updated in Google Calendar: {dto.Subject}");
                }
                else
                {
                    saved = await _calendarService.Events.Insert(googleEvent, TargetCalendarId).ExecuteAsync();
                    Log($"Event created in Google Calendar: {dto.Subject}");
                }

                // Occurrences moved or renamed in Outlook: override the matching instance of the Google series.
                foreach (var occurrence in dto.Modified)
                {
                    string instanceId = saved.Id + "_" + occurrence.InstanceKey;
                    try
                    {
                        var instance = await _calendarService.Events.Get(TargetCalendarId, instanceId).ExecuteAsync();
                        instance.Summary = occurrence.Subject;
                        instance.Location = occurrence.Location;
                        SetTimes(instance, dto.IsAllDay, occurrence.StartLocal, occurrence.EndLocal, occurrence.StartUtc, occurrence.EndUtc, dto.TimeZone);
                        await _calendarService.Events.Update(instance, TargetCalendarId, instanceId).ExecuteAsync();
                        Log($"Occurrence {occurrence.InstanceKey} of '{dto.Subject}' updated");
                    }
                    catch (Exception ex)
                    {
                        Log($"Occurrence {occurrence.InstanceKey} of '{dto.Subject}' not updated: {ex.Message}");
                    }
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

        private async Task DeleteGoogleEventAsync(string globalId, string subject)
        {
            await _syncLock.WaitAsync();
            try
            {
                if (!await ReadyAsync(subject))
                    return;

                var existingEvents = await FindGoogleEventsAsync(globalId);
                foreach (var existing in existingEvents)
                {
                    await _calendarService.Events.Delete(TargetCalendarId, existing.Id).ExecuteAsync();
                }
                Log(existingEvents.Count > 0
                    ? $"Event deleted (removed in Outlook): {subject}"
                    : $"Nothing to delete in Google Calendar for: {subject}");
            }
            catch (Exception ex)
            {
                Log($"Google Calendar API delete error for '{subject}': {ex.Message}");
            }
            finally
            {
                _syncLock.Release();
            }
        }

        // All-day events use local calendar dates (a UTC date can fall on the previous day east of Greenwich);
        // a series also carries its IANA zone so Google expands it across daylight saving changes.
        private static void SetTimes(Event target, bool allDay, DateTime startLocal, DateTime endLocal,
            DateTime startUtc, DateTime endUtc, string timeZone)
        {
            if (allDay)
            {
                target.Start = new EventDateTime { Date = startLocal.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) };
                target.End = new EventDateTime { Date = endLocal.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) };
            }
            else
            {
                target.Start = new EventDateTime
                {
                    DateTimeDateTimeOffset = new DateTimeOffset(DateTime.SpecifyKind(startUtc, DateTimeKind.Utc)),
                    TimeZone = timeZone
                };
                target.End = new EventDateTime
                {
                    DateTimeDateTimeOffset = new DateTimeOffset(DateTime.SpecifyKind(endUtc, DateTimeKind.Utc)),
                    TimeZone = timeZone
                };
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
            foreach (var folder in _calendarFolders)
            {
                try
                {
                    folder.BeforeItemMove -= CalendarFolder_BeforeItemMove;
                    Marshal.ReleaseComObject(folder);
                }
                catch { }
            }
            _calendarItems.Clear();
            _calendarFolders.Clear();
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
            public DateTime StartLocal { get; set; }
            public DateTime EndLocal { get; set; }
            public bool IsAllDay { get; set; }
            public string Organizer { get; set; }
            public bool IsCanceled { get; set; }
            public string TimeZone { get; set; }
            public IList<string> Recurrence { get; set; }
            public List<OccurrenceDto> Modified { get; } = new List<OccurrenceDto>();
        }

        private class OccurrenceDto
        {
            public string InstanceKey { get; set; }
            public string Subject { get; set; }
            public string Location { get; set; }
            public DateTime StartUtc { get; set; }
            public DateTime EndUtc { get; set; }
            public DateTime StartLocal { get; set; }
            public DateTime EndLocal { get; set; }
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
