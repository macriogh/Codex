using System;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAvailabilitySummarizer;

public sealed class MainForm : Form
{
    private readonly DateTimePicker _fromDatePicker;
    private readonly DateTimePicker _toDatePicker;
    private readonly DateTimePicker _windowStartPicker;
    private readonly DateTimePicker _windowEndPicker;
    private readonly Button _summarizeButton;
    private readonly TextBox _summaryText;

    public MainForm()
    {
        Text = "Outlook Availability Summarizer";
        Width = 920;
        Height = 700;
        StartPosition = FormStartPosition.CenterScreen;

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 160,
            ColumnCount = 4,
            RowCount = 3,
            Padding = new Padding(10),
            AutoSize = true,
        };

        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        panel.Controls.Add(new Label { Text = "From date", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 0);
        _fromDatePicker = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        panel.Controls.Add(_fromDatePicker, 1, 0);

        panel.Controls.Add(new Label { Text = "To date", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 2, 0);
        _toDatePicker = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(4) };
        panel.Controls.Add(_toDatePicker, 3, 0);

        panel.Controls.Add(new Label { Text = "Daily window starts", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, 1);
        _windowStartPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Value = DateTime.Today.AddHours(9),
        };
        panel.Controls.Add(_windowStartPicker, 1, 1);

        panel.Controls.Add(new Label { Text = "Daily window ends", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 2, 1);
        _windowEndPicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Time,
            ShowUpDown = true,
            Value = DateTime.Today.AddHours(17),
        };
        panel.Controls.Add(_windowEndPicker, 3, 1);

        _summarizeButton = new Button
        {
            Text = "Summarize Outlook Availability",
            Dock = DockStyle.Fill,
            Height = 35,
        };
        _summarizeButton.Click += SummarizeButton_Click;
        panel.Controls.Add(_summarizeButton, 0, 2);
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

    private void SummarizeButton_Click(object? sender, EventArgs e)
    {
        try
        {
            var request = BuildRequest();
            var summary = CreateAvailabilitySummary(request);
            _summaryText.Text = summary;
        }
        catch (System.Exception ex)
        {
            _summaryText.Text = $"Could not summarize Outlook availability: {ex.Message}";
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

    private static string CreateAvailabilitySummary(AvailabilityRequest request)
    {
        Outlook.Application? outlook = null;
        Outlook.NameSpace? session = null;
        Outlook.MAPIFolder? calendar = null;
        Outlook.Items? items = null;

        try
        {
            outlook = new Outlook.Application();
            session = outlook.GetNamespace("MAPI");
            calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            items = calendar.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");

            var filter = $"[Start] <= '{request.ToDate.AddDays(1):g}' AND [End] >= '{request.FromDate:g}'";
            var restricted = items.Restrict(filter);

            var summary = new StringBuilder();
            summary.AppendLine($"Availability summary from {request.FromDate:dddd, MMM d yyyy} to {request.ToDate:dddd, MMM d yyyy}.");
            summary.AppendLine($"Each day window: {DateTime.Today.Add(request.WindowStart):h:mm tt} - {DateTime.Today.Add(request.WindowEnd):h:mm tt}.");
            summary.AppendLine();

            var totalFree = TimeSpan.Zero;
            var totalBusy = TimeSpan.Zero;

            for (var date = request.FromDate; date <= request.ToDate; date = date.AddDays(1))
            {
                var dayStart = date.Add(request.WindowStart);
                var dayEnd = date.Add(request.WindowEnd);

                var busyIntervals = OutlookIntervalBuilder.GetBusyIntervals(restricted, dayStart, dayEnd);
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
        finally
        {
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

    private static string FormatIntervals(System.Collections.Generic.IReadOnlyList<TimeInterval> intervals)
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

public readonly record struct AvailabilityRequest(
    DateTime FromDate,
    DateTime ToDate,
    TimeSpan WindowStart,
    TimeSpan WindowEnd);

public readonly record struct TimeInterval(DateTime Start, DateTime End);

internal static class OutlookIntervalBuilder
{
    public static System.Collections.Generic.List<TimeInterval> GetBusyIntervals(Outlook.Items items, DateTime windowStart, DateTime windowEnd)
    {
        var busy = new System.Collections.Generic.List<TimeInterval>();

        for (var i = 1; i <= items.Count; i++)
        {
            if (items[i] is not Outlook.AppointmentItem appt)
            {
                continue;
            }

            var isBusy = appt.BusyStatus is Outlook.OlBusyStatus.olBusy or Outlook.OlBusyStatus.olTentative or Outlook.OlBusyStatus.olOutOfOffice;
            if (!isBusy)
            {
                continue;
            }

            var start = appt.Start;
            var end = appt.End;
            if (end <= windowStart || start >= windowEnd)
            {
                continue;
            }

            var clampedStart = start < windowStart ? windowStart : start;
            var clampedEnd = end > windowEnd ? windowEnd : end;
            if (clampedEnd > clampedStart)
            {
                busy.Add(new TimeInterval(clampedStart, clampedEnd));
            }
        }

        return busy;
    }

    public static System.Collections.Generic.List<TimeInterval> MergeIntervals(System.Collections.Generic.List<TimeInterval> intervals)
    {
        if (intervals.Count == 0)
        {
            return intervals;
        }

        intervals.Sort((a, b) => a.Start.CompareTo(b.Start));

        var merged = new System.Collections.Generic.List<TimeInterval> { intervals[0] };
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

    public static System.Collections.Generic.List<TimeInterval> GetFreeIntervals(DateTime windowStart, DateTime windowEnd, System.Collections.Generic.IReadOnlyList<TimeInterval> busy)
    {
        var free = new System.Collections.Generic.List<TimeInterval>();
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

    public static TimeSpan GetTotalDuration(System.Collections.Generic.IReadOnlyList<TimeInterval> intervals)
    {
        var total = TimeSpan.Zero;
        for (var i = 0; i < intervals.Count; i++)
        {
            total += intervals[i].End - intervals[i].Start;
        }

        return total;
    }
}
