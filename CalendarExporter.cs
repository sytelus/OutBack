using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public class CalendarExportOptions
    {
        public CalendarExportOptions()
        {
            Categories = new List<string>();
        }

        public string ExportFilePath { get; set; }
        public List<string> Categories { get; private set; }
        public bool FilterByCategories { get; set; }
        public bool IncludeUncategorized { get; set; }
        public bool AppointmentsOnly { get; set; }
        public bool MeetingsOnly { get; set; }
        public bool OrganizedByCurrentUserOnly { get; set; }
    }

    public class CalendarExportResult
    {
        public string FilePath { get; set; }
        public int TotalItems { get; set; }
        public int ExportedItems { get; set; }
        public int SkippedItems { get; set; }
        public int ErrorItems { get; set; }
    }

    public class CalendarExporter
    {
        private readonly string logFilePath;

        public CalendarExporter()
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".OutBack");
            Directory.CreateDirectory(logDirectory);
            logFilePath = Path.Combine(logDirectory, $"OutBack_CalendarExport_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        }

        public CalendarExportResult Export(Outlook.MAPIFolder calendarFolder, CalendarExportOptions options)
        {
            if (calendarFolder.DefaultItemType != Outlook.OlItemType.olAppointmentItem)
                throw new InvalidOperationException("The selected folder is not a Calendar folder.");

            if (options == null || string.IsNullOrWhiteSpace(options.ExportFilePath))
                throw new ArgumentException("An export file path is required.", "options");

            var result = new CalendarExportResult { FilePath = options.ExportFilePath };
            var timeZoneComponents = new HashSet<string>(StringComparer.Ordinal);
            var eventComponents = new List<string>();
            var selectedCategories = new HashSet<string>(options.Categories, StringComparer.OrdinalIgnoreCase);
            HashSet<string> currentUserIdentities = options.OrganizedByCurrentUserOnly
                ? GetCurrentUserIdentities()
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Outlook.Items items = null;
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                items = calendarFolder.Items;
                try
                {
                    items.Sort("[Start]");
                }
                catch
                {
                    // Sorting is only for deterministic export order.
                }

                result.TotalItems = items.Count;
                Log($"Calendar export started: Source={calendarFolder.FolderPath}, Items={result.TotalItems}, File={options.ExportFilePath}");

                for (int i = 1; i <= result.TotalItems; i++)
                {
                    object item = null;

                    try
                    {
                        item = items[i];
                        Outlook.AppointmentItem appointment = null;
                        if (!TryCastItem(item, out appointment))
                        {
                            result.SkippedItems++;
                            continue;
                        }

                        if (!ShouldExport(appointment, options, selectedCategories, currentUserIdentities))
                        {
                            result.SkippedItems++;
                            continue;
                        }

                        ExportAppointment(appointment, timeZoneComponents, eventComponents);
                        result.ExportedItems++;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorItems++;
                        Log($"Error exporting calendar item {i}", ex);
                    }
                    finally
                    {
                        if (item != null) Marshal.ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
            }

            WriteCalendarFile(options.ExportFilePath, timeZoneComponents, eventComponents);

            stopwatch.Stop();
            Log($"Calendar export completed: Exported={result.ExportedItems}, Skipped={result.SkippedItems}, Errors={result.ErrorItems}, Time={stopwatch.Elapsed}");
            return result;
        }

        private void ExportAppointment(Outlook.AppointmentItem appointment, HashSet<string> timeZoneComponents, List<string> eventComponents)
        {
            try
            {
                ExportAppointmentUsingOutlook(appointment, timeZoneComponents, eventComponents);
            }
            catch (Exception ex)
            {
                Log("Outlook iCalendar save failed; using basic event fallback.", ex);
                eventComponents.Add(BuildBasicVEvent(appointment));
            }
        }

        private static void ExportAppointmentUsingOutlook(Outlook.AppointmentItem appointment, HashSet<string> timeZoneComponents, List<string> eventComponents)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"OutBack_{Guid.NewGuid():N}.ics");

            try
            {
                appointment.SaveAs(tempPath, Outlook.OlSaveAsType.olICal);
                string calendarText = File.ReadAllText(tempPath);
                bool foundEvent = false;

                foreach (string timeZoneComponent in ExtractComponents(calendarText, "VTIMEZONE"))
                {
                    timeZoneComponents.Add(timeZoneComponent);
                }

                foreach (string eventComponent in ExtractComponents(calendarText, "VEVENT"))
                {
                    eventComponents.Add(eventComponent);
                    foundEvent = true;
                }

                if (!foundEvent)
                    throw new InvalidOperationException("Outlook did not write a VEVENT component.");
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch
                {
                }
            }
        }

        private static IEnumerable<string> ExtractComponents(string calendarText, string componentName)
        {
            string begin = "BEGIN:" + componentName;
            string end = "END:" + componentName;
            var lines = new List<string>();
            bool isCapturing = false;

            using (StringReader reader = new StringReader(calendarText))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string trimmed = line.Trim();
                    if (!isCapturing && string.Equals(trimmed, begin, StringComparison.OrdinalIgnoreCase))
                    {
                        isCapturing = true;
                        lines.Clear();
                    }

                    if (isCapturing)
                    {
                        lines.Add(line);
                        if (string.Equals(trimmed, end, StringComparison.OrdinalIgnoreCase))
                        {
                            yield return string.Join("\r\n", lines);
                            isCapturing = false;
                            lines.Clear();
                        }
                    }
                }
            }
        }

        private static void WriteCalendarFile(string exportFilePath, HashSet<string> timeZoneComponents, List<string> eventComponents)
        {
            using (var writer = new StreamWriter(exportFilePath, false, new UTF8Encoding(false)))
            {
                writer.WriteLine("BEGIN:VCALENDAR");
                writer.WriteLine("PRODID:-//OutBack//Calendar Export//EN");
                writer.WriteLine("VERSION:2.0");
                writer.WriteLine("CALSCALE:GREGORIAN");
                writer.WriteLine("METHOD:PUBLISH");

                foreach (string timeZoneComponent in timeZoneComponents)
                {
                    writer.WriteLine(timeZoneComponent);
                }

                foreach (string eventComponent in eventComponents)
                {
                    writer.WriteLine(eventComponent);
                }

                writer.WriteLine("END:VCALENDAR");
            }
        }

        private static bool ShouldExport(Outlook.AppointmentItem appointment, CalendarExportOptions options, HashSet<string> selectedCategories, HashSet<string> currentUserIdentities)
        {
            bool isMeeting = IsMeeting(appointment);

            if (options.AppointmentsOnly && isMeeting)
                return false;

            if (options.MeetingsOnly && !isMeeting)
                return false;

            if (options.OrganizedByCurrentUserOnly && !IsOrganizedByCurrentUser(appointment, currentUserIdentities))
                return false;

            if (!options.FilterByCategories)
                return true;

            List<string> itemCategories = SplitCategories(SafeGetString(() => appointment.Categories));
            if (itemCategories.Count == 0)
                return options.IncludeUncategorized;

            foreach (string category in itemCategories)
            {
                if (selectedCategories.Contains(category))
                    return true;
            }

            return false;
        }

        private static bool IsMeeting(Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsOrganizedByCurrentUser(Outlook.AppointmentItem appointment, HashSet<string> currentUserIdentities)
        {
            var itemIdentities = new List<string>
            {
                SafeGetString(() => appointment.Organizer)
            };

            Outlook.PropertyAccessor propertyAccessor = null;
            try
            {
                propertyAccessor = appointment.PropertyAccessor;
                AddPropertyValue(itemIdentities, propertyAccessor, "http://schemas.microsoft.com/mapi/proptag/0x0042001E");
                AddPropertyValue(itemIdentities, propertyAccessor, "http://schemas.microsoft.com/mapi/proptag/0x0065001E");
                AddPropertyValue(itemIdentities, propertyAccessor, "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E");
                AddPropertyValue(itemIdentities, propertyAccessor, "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E");
                AddPropertyValue(itemIdentities, propertyAccessor, "http://schemas.microsoft.com/mapi/proptag/0x3FF8001E");
            }
            catch
            {
            }
            finally
            {
                if (propertyAccessor != null) Marshal.ReleaseComObject(propertyAccessor);
            }

            foreach (string itemIdentity in itemIdentities)
            {
                if (currentUserIdentities.Contains(NormalizeIdentity(itemIdentity)))
                    return true;
            }

            return false;
        }

        private static HashSet<string> GetCurrentUserIdentities()
        {
            var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry addressEntry = null;
            Outlook.ExchangeUser exchangeUser = null;

            try
            {
                currentUser = Globals.ThisAddIn.Application.Session.CurrentUser;
                AddIdentity(identities, SafeGetString(() => currentUser.Name));
                AddIdentity(identities, SafeGetString(() => currentUser.Address));

                addressEntry = currentUser.AddressEntry;
                if (addressEntry != null)
                {
                    AddIdentity(identities, SafeGetString(() => addressEntry.Name));
                    AddIdentity(identities, SafeGetString(() => addressEntry.Address));

                    try
                    {
                        exchangeUser = addressEntry.GetExchangeUser();
                        if (exchangeUser != null)
                            AddIdentity(identities, SafeGetString(() => exchangeUser.PrimarySmtpAddress));
                    }
                    catch
                    {
                    }
                }
            }
            finally
            {
                if (exchangeUser != null) Marshal.ReleaseComObject(exchangeUser);
                if (addressEntry != null) Marshal.ReleaseComObject(addressEntry);
                if (currentUser != null) Marshal.ReleaseComObject(currentUser);
            }

            return identities;
        }

        private static void AddPropertyValue(List<string> values, Outlook.PropertyAccessor propertyAccessor, string schemaName)
        {
            try
            {
                object value = propertyAccessor.GetProperty(schemaName);
                if (value != null)
                    values.Add(value.ToString());
            }
            catch
            {
            }
        }

        private static void AddIdentity(HashSet<string> identities, string identity)
        {
            string normalized = NormalizeIdentity(identity);
            if (!string.IsNullOrEmpty(normalized))
                identities.Add(normalized);
        }

        private static string NormalizeIdentity(string identity)
        {
            return string.IsNullOrWhiteSpace(identity) ? string.Empty : identity.Trim().ToLowerInvariant();
        }

        private static string BuildBasicVEvent(Outlook.AppointmentItem appointment)
        {
            var builder = new StringBuilder();
            string uid = SafeGetString(() => appointment.GlobalAppointmentID);
            if (string.IsNullOrWhiteSpace(uid))
                uid = SafeGetString(() => appointment.EntryID);
            if (string.IsNullOrWhiteSpace(uid))
                uid = Guid.NewGuid().ToString("N");

            DateTime start = SafeGetDateTime(() => appointment.Start);
            DateTime end = SafeGetDateTime(() => appointment.End);
            bool allDay = SafeGetBool(() => appointment.AllDayEvent);

            AppendFoldedLine(builder, "BEGIN:VEVENT");
            AppendFoldedLine(builder, "UID:" + EscapeIcsText(uid));
            AppendFoldedLine(builder, "DTSTAMP:" + ToUtcIcs(DateTime.Now));

            if (allDay)
            {
                DateTime startDate = start.Date;
                DateTime endDate = end.Date;
                if (endDate <= startDate)
                    endDate = startDate.AddDays(1);

                AppendFoldedLine(builder, "DTSTART;VALUE=DATE:" + startDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
                AppendFoldedLine(builder, "DTEND;VALUE=DATE:" + endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
            }
            else
            {
                AppendFoldedLine(builder, "DTSTART:" + ToUtcIcs(start));
                AppendFoldedLine(builder, "DTEND:" + ToUtcIcs(end));
            }

            AppendProperty(builder, "SUMMARY", SafeGetString(() => appointment.Subject));
            AppendProperty(builder, "DESCRIPTION", SafeGetString(() => appointment.Body));
            AppendProperty(builder, "LOCATION", SafeGetString(() => appointment.Location));
            AppendProperty(builder, "CATEGORIES", SafeGetString(() => appointment.Categories));
            AppendFoldedLine(builder, "END:VEVENT");

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private static void AppendProperty(StringBuilder builder, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
                return;

            AppendFoldedLine(builder, name + ":" + EscapeIcsText(value));
        }

        private static void AppendFoldedLine(StringBuilder builder, string line)
        {
            const int maxLength = 75;
            string remaining = line;

            while (remaining.Length > maxLength)
            {
                builder.Append(remaining.Substring(0, maxLength));
                builder.Append("\r\n");
                remaining = " " + remaining.Substring(maxLength);
            }

            builder.Append(remaining);
            builder.Append("\r\n");
        }

        private static string EscapeIcsText(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value
                .Replace("\\", "\\\\")
                .Replace(";", "\\;")
                .Replace(",", "\\,")
                .Replace("\r\n", "\\n")
                .Replace("\n", "\\n")
                .Replace("\r", "\\n");
        }

        private static string ToUtcIcs(DateTime value)
        {
            return value.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
        }

        private static List<string> SplitCategories(string categories)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(categories))
                return result;

            foreach (string category in categories.Split(','))
            {
                string trimmed = category.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    result.Add(trimmed);
            }

            return result;
        }

        private static string SafeGetString(Func<string> getter)
        {
            try
            {
                return getter() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime SafeGetDateTime(Func<DateTime> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return DateTime.Now;
            }
        }

        private static bool SafeGetBool(Func<bool> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return false;
            }
        }

        private static bool TryCastItem<T>(object item, out T typedItem) where T : class
        {
            typedItem = null;
            if (item == null)
                return false;

            try
            {
                typedItem = item as T;
                return typedItem != null;
            }
            catch (InvalidCastException)
            {
                return false;
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x80004002)
                    return false;

                throw;
            }
        }

        private void Log(string message, Exception ex = null)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            if (ex != null)
            {
                logEntry += $"\nException: {ex.GetType().FullName}" +
                           $"\nMessage: {ex.Message}" +
                           $"\nStack Trace:\n{ex.StackTrace}";
            }

            File.AppendAllText(logFilePath, logEntry + Environment.NewLine + Environment.NewLine);
        }
    }
}
