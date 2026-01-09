using System;
using System.Threading.Tasks;
using ExcelDna.Integration;

public static class AsyncFunctions
{
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

