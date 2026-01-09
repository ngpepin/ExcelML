using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

public class DlRibbon : ExcelRibbon
{
    private IRibbonUI _ribbon;

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

    public void OnLoad(IRibbonUI ribbonUi) => _ribbon = ribbonUi;

    public void OnHelloClick(IRibbonControl control)
        => MessageBox.Show("Hello from an Excel-DNA Ribbon button!");

    public void OnInvalidateClick(IRibbonControl control)
        => _ribbon?.Invalidate();
}
