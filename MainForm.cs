using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Azure.Identity;
using Microsoft.Graph;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAvailabilitySummarizer;

public sealed class MainForm : Form
{
    private readonly DateTimePicker _fromDatePicker;
    private readonly DateTimePicker _toDatePicker;
    private readonly DateTimePicker _windowStartPicker;
    private readonly DateTimePicker _windowEndPicker;
    private readonly ComboBox _sourceCombo;
    private readonly TextBox _tenantIdText;
    private readonly TextBox _clientIdText;
    private readonly Button _summarizeButton;
    private readonly TextBox _summaryText;

    public MainForm()
    {
        Text = "Outlook Availability Summarizer";
        Width = 980;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 255,
            ColumnCount = 4,
            RowCount = 5,
            Padding = new Padding(10),
            AutoSize = true,
        };

        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        panel.Controls.Add(new Label { Text = "Calendar source", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 0);
        _sourceCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        _sourceCombo.Items.Add("Classic Outlook (COM)");
        _sourceCombo.Items.Add("New Outlook / Microsoft 365 (Graph API)");
        _sourceCombo.SelectedIndex = 0;
        panel.Controls.Add(_sourceCombo, 1, 0);
        panel.SetColumnSpan(_sourceCombo, 3);

        panel.Controls.Add(new Label { Text = "From date", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 1);
        _fromDatePicker = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        panel.Controls.Add(_fromDatePicker, 1, 1);

        panel.Controls.Add(new Label { Text = "To date", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 2, 1);
        _toDatePicker = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(4) };
        panel.Controls.Add(_toDatePicker, 3, 1);

        panel.Controls.Add(new Label { Text = "Daily window starts", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 2);
        _windowStartPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Value = DateTime.Today.AddHours(9),
        };
        panel.Controls.Add(_windowStartPicker, 1, 2);

        panel.Controls.Add(new Label { Text = "Daily window ends", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 2, 2);
        _windowEndPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Value = DateTime.Today.AddHours(17),
        };
        panel.Controls.Add(_windowEndPicker, 3, 2);

        panel.Controls.Add(new Label { Text = "Tenant ID (Graph)", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 3);
        _tenantIdText = new TextBox { Dock = DockStyle.Fill, Text = "common" };
        panel.Controls.Add(_tenantIdText, 1, 3);

        panel.Controls.Add(new Label { Text = "Client ID (Graph app)", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 2, 3);
        _clientIdText = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "Enter Azure AD app Client ID" };
        panel.Controls.Add(_clientIdText, 3, 3);

        _summarizeButton = new Button
        {
            Text = "Summarize Outlook Availability",
            Dock = DockStyle.Fill,
            Height = 35,
        };
        _summarizeButton.Click += SummarizeButton_Click;
        panel.Controls.Add(_summarizeButton, 0, 4);
        panel.SetColumnSpan(_summarizeButton, 4);

        _summaryText = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new System.Drawing.Font("Segoe UI", 10),
            ReadOnly = true,
        };

        Controls.Add(_summaryText);
        Controls.Add(panel);

        panel.Dock = DockStyle.Top;
        _summaryText.BringToFront();
    }

    private async void SummarizeButton_Click(object? sender, EventArgs e)
    {
        _summarizeButton.Enabled = false;
        try
        {
            var request = BuildRequest();
            var provider = BuildProvider();
            _summaryText.Text = "Collecting calendar data...";
            var summary = await CreateAvailabilitySummaryAsync(request, provider);
            _summaryText.Text = summary;
        }
        catch (System.Exception ex)
        {
            _summaryText.Text = $"Could not summarize Outlook availability: {ex.Message}";
        }
        finally
        {
            _summarizeButton.Enabled = true;
        }
    }

    private AvailabilityRequest BuildRequest()
    {
        var fromDate = _fromDatePicker.Value.Date;
        var toDate = _toDatePicker.Value.Date;
        if (toDate < fromDate)
        {
            throw new InvalidOperationException("The end date must be on or after the start date.");
        }

        var startTime = _windowStartPicker.Value.TimeOfDay;
        var endTime = _windowEndPicker.Value.TimeOfDay;
        if (endTime <= startTime)
        {
            throw new InvalidOperationException("Daily end time must be after the daily start time.");
        }

        return new AvailabilityRequest(fromDate, toDate, startTime, endTime);
    }

    private ICalendarProvider BuildProvider()
    {
        if (_sourceCombo.SelectedIndex == 0)
        {
            return new OutlookComCalendarProvider();
        }

        if (string.IsNullOrWhiteSpace(_clientIdText.Text))
        {
            throw new InvalidOperationException("Client ID is required for Graph mode (new Outlook). Create an Entra app with delegated Calendars.Read permission.");
        }

        return new GraphCalendarProvider(_tenantIdText.Text.Trim(), _clientIdText.Text.Trim());
    }

    private static async Task<string> CreateAvailabilitySummaryAsync(AvailabilityRequest request, ICalendarProvider provider)
    {
        var summary = new StringBuilder();
        summary.AppendLine($"Availability summary from {request.FromDate:dddd, MMM d yyyy} to {request.ToDate:dddd, MMM d yyyy}.");
        summary.AppendLine($"Each day window: {DateTime.Today.Add(request.WindowStart):h:mm tt} - {DateTime.Today.Add(request.WindowEnd):h:mm tt}.");
        summary.AppendLine();

        var allEvents = await provider.GetAppointmentsAsync(request.FromDate, request.ToDate.AddDays(1));

        var totalFree = TimeSpan.Zero;
        var totalBusy = TimeSpan.Zero;

        for (var date = request.FromDate; date <= request.ToDate; date = date.AddDays(1))
        {
            var dayStart = date.Add(request.WindowStart);
            var dayEnd = date.Add(request.WindowEnd);

            var busyIntervals = OutlookIntervalBuilder.GetBusyIntervals(allEvents, dayStart, dayEnd);
            var normalizedBusy = OutlookIntervalBuilder.MergeIntervals(busyIntervals);
            var freeIntervals = OutlookIntervalBuilder.GetFreeIntervals(dayStart, dayEnd, normalizedBusy);

            var dayBusy = OutlookIntervalBuilder.GetTotalDuration(normalizedBusy);
            var dayFree = OutlookIntervalBuilder.GetTotalDuration(freeIntervals);
            totalBusy += dayBusy;
            totalFree += dayFree;

            summary.AppendLine($"{date:dddd, MMM d}: free {FormatDuration(dayFree)}, busy {FormatDuration(dayBusy)}.");
            if (freeIntervals.Count == 0)
            {
                summary.AppendLine("  No free blocks in the selected window.");
            }
            else
            {
                summary.AppendLine($"  Free blocks: {FormatIntervals(freeIntervals)}");
            }
        }

        summary.AppendLine();
        summary.AppendLine($"Overall free time: {FormatDuration(totalFree)}.");
        summary.AppendLine($"Overall busy time: {FormatDuration(totalBusy)}.");

        return summary.ToString();
    }

    private static string FormatIntervals(IReadOnlyList<TimeInterval> intervals)
    {
        var parts = new string[intervals.Count];
        for (var i = 0; i < intervals.Count; i++)
        {
            var interval = intervals[i];
            parts[i] = $"{interval.Start:h:mm tt}-{interval.End:h:mm tt}";
        }

        return string.Join(", ", parts);
    }

    private static string FormatDuration(TimeSpan span)
    {
        if (span.TotalMinutes < 1)
        {
            return "0 minutes";
        }

        var hours = (int)span.TotalHours;
        var minutes = span.Minutes;

        if (hours > 0 && minutes > 0)
        {
            return $"{hours}h {minutes}m";
        }

        if (hours > 0)
        {
            return hours == 1 ? "1 hour" : $"{hours} hours";
        }

        return minutes == 1 ? "1 minute" : $"{minutes} minutes";
    }
}

public readonly record struct AvailabilityRequest(DateTime FromDate, DateTime ToDate, TimeSpan WindowStart, TimeSpan WindowEnd);

public readonly record struct TimeInterval(DateTime Start, DateTime End);

public readonly record struct CalendarAppointment(DateTime Start, DateTime End, bool IsBusy);

internal interface ICalendarProvider
{
    Task<List<CalendarAppointment>> GetAppointmentsAsync(DateTime rangeStart, DateTime rangeEndExclusive);
}

internal sealed class OutlookComCalendarProvider : ICalendarProvider
{
    public Task<List<CalendarAppointment>> GetAppointmentsAsync(DateTime rangeStart, DateTime rangeEndExclusive)
    {
        Outlook.Application? outlook = null;
        Outlook.NameSpace? session = null;
        Outlook.MAPIFolder? calendar = null;
        Outlook.Items? items = null;
        Outlook.Items? restricted = null;

        try
        {
            outlook = new Outlook.Application();
            session = outlook.GetNamespace("MAPI");
            calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            items = calendar.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");

            var filter = $"[Start] <= '{rangeEndExclusive:g}' AND [End] >= '{rangeStart:g}'";
            restricted = items.Restrict(filter);

            var results = new List<CalendarAppointment>();
            for (var i = 1; i <= restricted.Count; i++)
            {
                if (restricted[i] is not Outlook.AppointmentItem appt)
                {
                    continue;
                }

                var isBusy = appt.BusyStatus is Outlook.OlBusyStatus.olBusy or Outlook.OlBusyStatus.olTentative or Outlook.OlBusyStatus.olOutOfOffice;
                results.Add(new CalendarAppointment(appt.Start, appt.End, isBusy));
            }

            return Task.FromResult(results);
        }
        catch (System.Exception ex)
        {
            throw new InvalidOperationException("Classic Outlook COM access failed. If you are using new Outlook v1.2026.x, switch source to Graph API.", ex);
        }
        finally
        {
            if (restricted is not null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
            }
            if (items is not null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            if (calendar is not null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
            }
            if (session is not null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
            }
            if (outlook is not null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outlook);
            }
        }
    }
}

internal sealed class GraphCalendarProvider : ICalendarProvider
{
    private static readonly string[] Scopes = ["Calendars.Read", "offline_access", "openid", "profile"];
    private readonly string _tenantId;
    private readonly string _clientId;

    public GraphCalendarProvider(string tenantId, string clientId)
    {
        _tenantId = string.IsNullOrWhiteSpace(tenantId) ? "common" : tenantId;
        _clientId = clientId;
    }

    public async Task<List<CalendarAppointment>> GetAppointmentsAsync(DateTime rangeStart, DateTime rangeEndExclusive)
    {
        var credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            TenantId = _tenantId,
            ClientId = _clientId,
            DeviceCodeCallback = (callback, cancellationToken) =>
            {
                MessageBox.Show(callback.Message, "Microsoft sign-in required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return Task.CompletedTask;
            },
        });

        var graphClient = new GraphServiceClient(credential, Scopes);
        var events = await graphClient.Me.CalendarView.GetAsync(config =>
        {
            config.QueryParameters.StartDateTime = rangeStart.ToString("o");
            config.QueryParameters.EndDateTime = rangeEndExclusive.ToString("o");
            config.QueryParameters.Top = 1000;
            config.Headers.Add("Prefer", "outlook.timezone=\"UTC\"");
        });

        var results = new List<CalendarAppointment>();
        if (events?.Value is null)
        {
            return results;
        }

        foreach (var item in events.Value)
        {
            if (item.Start?.DateTime is null || item.End?.DateTime is null)
            {
                continue;
            }

            if (!DateTime.TryParse(item.Start.DateTime, out var start) || !DateTime.TryParse(item.End.DateTime, out var end))
            {
                continue;
            }

            var showAs = item.ShowAs?.ToString() ?? string.Empty;
            var isBusy = showAs.Contains("busy", StringComparison.OrdinalIgnoreCase)
                || showAs.Contains("tentative", StringComparison.OrdinalIgnoreCase)
                || showAs.Contains("oof", StringComparison.OrdinalIgnoreCase);

            results.Add(new CalendarAppointment(start.ToLocalTime(), end.ToLocalTime(), isBusy));
        }

        return results;
    }
}

internal static class OutlookIntervalBuilder
{
    public static List<TimeInterval> GetBusyIntervals(IReadOnlyList<CalendarAppointment> appointments, DateTime windowStart, DateTime windowEnd)
    {
        var busy = new List<TimeInterval>();

        for (var i = 0; i < appointments.Count; i++)
        {
            var appt = appointments[i];
            if (!appt.IsBusy)
            {
                continue;
            }

            if (appt.End <= windowStart || appt.Start >= windowEnd)
            {
                continue;
            }

            var clampedStart = appt.Start < windowStart ? windowStart : appt.Start;
            var clampedEnd = appt.End > windowEnd ? windowEnd : appt.End;
            if (clampedEnd > clampedStart)
            {
                busy.Add(new TimeInterval(clampedStart, clampedEnd));
            }
        }

        return busy;
    }

    public static List<TimeInterval> MergeIntervals(List<TimeInterval> intervals)
    {
        if (intervals.Count == 0)
        {
            return intervals;
        }

        intervals.Sort((a, b) => a.Start.CompareTo(b.Start));

        var merged = new List<TimeInterval> { intervals[0] };
        for (var i = 1; i < intervals.Count; i++)
        {
            var current = intervals[i];
            var last = merged[^1];

            if (current.Start <= last.End)
            {
                var mergedInterval = new TimeInterval(last.Start, current.End > last.End ? current.End : last.End);
                merged[^1] = mergedInterval;
            }
            else
            {
                merged.Add(current);
            }
        }

        return merged;
    }

    public static List<TimeInterval> GetFreeIntervals(DateTime windowStart, DateTime windowEnd, IReadOnlyList<TimeInterval> busy)
    {
        var free = new List<TimeInterval>();
        var cursor = windowStart;

        for (var i = 0; i < busy.Count; i++)
        {
            var busyInterval = busy[i];
            if (busyInterval.Start > cursor)
            {
                free.Add(new TimeInterval(cursor, busyInterval.Start));
            }

            if (busyInterval.End > cursor)
            {
                cursor = busyInterval.End;
            }
        }

        if (cursor < windowEnd)
        {
            free.Add(new TimeInterval(cursor, windowEnd));
        }

        return free;
    }

    public static TimeSpan GetTotalDuration(IReadOnlyList<TimeInterval> intervals)
    {
        var total = TimeSpan.Zero;
        for (var i = 0; i < intervals.Count; i++)
        {
            total += intervals[i].End - intervals[i].Start;
        }

        return total;
    }
}
