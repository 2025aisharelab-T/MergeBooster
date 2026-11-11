using System;
using System.Diagnostics;

namespace ExcelDiffAggregator
{
    // percent: 0..100, status: what is being worked on right now
    public delegate void ProgressHandler(int percent, string status);

    // Turns raw progress numbers into a calm stream of UI updates:
    // clamped to 0..100, never moving backwards, whole percents only, and throttled.
    // Call it from one thread only.
    public sealed class ProgressTracker
    {
        private readonly ProgressHandler handler;
        private readonly int minIntervalMs;
        private readonly Func<long> clockMs;

        private double highest;
        private bool hasSent;
        private int lastSentPercent;
        private string lastSentStatus;
        private long lastSentAt;

        public ProgressTracker(ProgressHandler handler, int minIntervalMs, Func<long> clockMs)
        {
            this.handler = handler;
            this.minIntervalMs = minIntervalMs;
            this.clockMs = clockMs ?? StopwatchMs;
        }

        private static long StopwatchMs()
        {
            return Stopwatch.GetTimestamp() * 1000L / Stopwatch.Frequency;
        }

        // Highest percent reported so far, floored (100 is only reached when the caller says so)
        public int Current
        {
            get { return (int)Math.Floor(highest); }
        }

        // True when a non-forced report would go out now. Callers check this before building a status string.
        public bool Due
        {
            get
            {
                if (handler == null) return false;
                return !hasSent || clockMs() - lastSentAt >= minIntervalMs;
            }
        }

        // force: send even inside the throttle interval (phase changes, completion)
        public void Report(double percent, string status, bool force)
        {
            if (handler == null) return;

            if (double.IsNaN(percent) || percent < 0) percent = 0;
            if (percent > 100) percent = 100;
            if (percent > highest) highest = percent;

            int shown = Current;
            string text = status ?? "";
            if (hasSent && shown == lastSentPercent && text == lastSentStatus) return;

            long now = clockMs();
            if (!force && hasSent && now - lastSentAt < minIntervalMs) return;

            hasSent = true;
            lastSentAt = now;
            lastSentPercent = shown;
            lastSentStatus = text;
            handler(shown, text);
        }
    }
}
