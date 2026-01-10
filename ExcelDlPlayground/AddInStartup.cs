using ExcelDna.Integration;

/// <summary>
/// Excel-DNA add-in bootstrap that triggers a one-time recalc when the add-in loads.
/// </summary>
public class AddInStartup : IExcelAddIn
{
    /// <summary>
    /// Invoked by Excel when the add-in is opened; issues a calculate-now to refresh UDF results once.
    /// </summary>
    public void AutoOpen()
    {
        try
        {
            // Recalculate on add-in load so DL.MODEL_CREATE cells refresh once when the workbook opens.
            XlCall.Excel(XlCall.xlcCalculateNow);
        }
        catch
        {
        }
    }

    /// <summary>
    /// Invoked by Excel when the add-in is unloaded; no cleanup required for this add-in.
    /// </summary>
    public void AutoClose()
    {
        // no-op
    }
}
