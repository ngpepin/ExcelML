using ExcelDna.Integration;

public class AddInStartup : IExcelAddIn
{
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

    public void AutoClose()
    {
        // no-op
    }
}
