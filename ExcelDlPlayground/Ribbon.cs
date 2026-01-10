using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

/// <summary>
/// Excel-DNA ribbon definition for the deep learning playground, exposing simple UI buttons.
/// </summary>
public class DlRibbon : ExcelRibbon
{
    private IRibbonUI _ribbon;

    /// <summary>
    /// Returns the ribbon XML to register custom tab and buttons with Excel.
    /// </summary>
    /// <param name="ribbonId">Unused ribbon identifier provided by Excel.</param>
    /// <returns>Ribbon XML string.</returns>
    public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='dlTab' label='Deep Learning'>
        <group id='dlGroup' label='Playground'>
          <button id='btnHello' label='Hello' size='large' imageMso='HappyFace'
                  onAction='OnHelloClick' />
          <button id='btnInvalidate' label='Refresh UI' imageMso='RefreshAll'
                  onAction='OnInvalidateClick' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    /// <summary>
    /// Captures the Excel ribbon instance when the UI loads so callbacks can invalidate later.
    /// </summary>
    /// <param name="ribbonUi">Ribbon interface provided by Excel.</param>
    public void OnLoad(IRibbonUI ribbonUi) => _ribbon = ribbonUi;

    /// <summary>
    /// Displays a simple message box to confirm ribbon callbacks are wired.
    /// </summary>
    /// <param name="control">Ribbon control invoking the action.</param>
    public void OnHelloClick(IRibbonControl control)
        => MessageBox.Show("Hello from an Excel-DNA Ribbon button!");

    /// <summary>
    /// Forces the ribbon to refresh its UI state.
    /// </summary>
    /// <param name="control">Ribbon control invoking the action.</param>
    public void OnInvalidateClick(IRibbonControl control)
        => _ribbon?.Invalidate();
}
