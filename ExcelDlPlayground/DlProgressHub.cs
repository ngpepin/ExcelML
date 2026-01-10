using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using ExcelDna.Integration;

internal static class DlProgressHub
{
    // modelId -> observers
    private static readonly ConcurrentDictionary<string, HashSet<IExcelObserver>> _subs =
        new ConcurrentDictionary<string, HashSet<IExcelObserver>>(StringComparer.OrdinalIgnoreCase);

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

    public static void Publish(string modelId)
    {
        if (!_subs.TryGetValue(modelId, out var set)) return;

        IExcelObserver[] observers;
        lock (set) observers = new List<IExcelObserver>(set).ToArray();

        foreach (var obs in observers)
        {
            try { obs.OnNext(modelId); } catch { }
        }
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
