using System;
using System.Threading.Tasks;
using ExcelDna.Integration;

/// <summary>
/// Async-friendly Excel functions demonstrating background work via ExcelAsyncUtil.
/// </summary>
public static class AsyncFunctions
{
    /// <summary>
    /// Asynchronously waits for the specified milliseconds and returns a timestamp when complete.
    /// </summary>
    /// <param name="ms">Delay duration in milliseconds.</param>
    /// <returns>Timestamp string indicating when the wait finished.</returns>
    [ExcelFunction(Description = "Async wait (ms) and return a timestamp string")]
    public static object WaitAsync(int ms)
    {
        return ExcelAsyncUtil.RunTask(
            nameof(WaitAsync),
            new object[] { ms },
            async () =>
            {
                await Task.Delay(ms).ConfigureAwait(false);
                return $"Done at {DateTime.Now:HH:mm:ss.fff}";
            });
    }
}

