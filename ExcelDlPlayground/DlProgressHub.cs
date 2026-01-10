using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using ExcelDna.Integration;

internal static class DlProgressHub
{
    // modelId -> observers
    private static readonly ConcurrentDictionary<string, HashSet<IExcelObserver>> _subs =
        new ConcurrentDictionary<string, HashSet<IExcelObserver>>(StringComparer.OrdinalIgnoreCase);

    // modelId -> queued flag (prevents flooding Excel with macros)
    private static readonly ConcurrentDictionary<string, byte> _pending =
        new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

    public static IDisposable Subscribe(string modelId, IExcelObserver observer)
    {
        var set = _subs.GetOrAdd(modelId, _ => new HashSet<IExcelObserver>());
        lock (set) set.Add(observer);

        return new ActionDisposable(() =>
        {
            if (_subs.TryGetValue(modelId, out var s))
            {
                lock (s) s.Remove(observer);
            }
        });
    }

    /// <summary>
    /// Publish an update to all observers of modelId.
    /// IMPORTANT: This marshals callbacks onto Excel's main thread.
    /// </summary>
    public static void Publish(string modelId)
    {
        if (string.IsNullOrWhiteSpace(modelId))
            return;

        // Coalesce: if already queued, don't queue again.
        if (!_pending.TryAdd(modelId, 0))
            return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                if (!_subs.TryGetValue(modelId, out var set))
                    return;

                IExcelObserver[] observers;
                lock (set) observers = new List<IExcelObserver>(set).ToArray();

                foreach (var obs in observers)
                {
                    try
                    {
                        // value payload doesn't matter for your WrappedObserver pattern
                        obs.OnNext(modelId);
                    }
                    catch
                    {
                        // swallow to keep Excel stable
                    }
                }
            }
            finally
            {
                _pending.TryRemove(modelId, out _);
            }
        });
    }

    private sealed class ActionDisposable : IDisposable
    {
        private readonly Action _dispose;
        private bool _done;
        public ActionDisposable(Action dispose) => _dispose = dispose;

        public void Dispose()
        {
            if (_done) return;
            _done = true;
            _dispose();
        }
    }
}
