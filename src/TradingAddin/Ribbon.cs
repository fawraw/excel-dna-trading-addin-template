using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace TradingAddin;

[ComVisible(true)]
public class Ribbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        // Minimal ribbon with a "Send selection" button and a "Diagnostics" button.
        return """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
          <ribbon>
            <tabs>
              <tab id="tabTrading" label="Trading">
                <group id="grpCapture" label="Capture">
                  <button id="btnSendSelection"
                          label="Send selection"
                          imageMso="GroupSendForReview"
                          size="large"
                          onAction="OnSendSelection" />
                  <button id="btnToggleWatcher"
                          label="Toggle cell watcher"
                          imageMso="ControlEvent"
                          size="large"
                          onAction="OnToggleWatcher" />
                </group>
                <group id="grpDiag" label="Diagnostics">
                  <button id="btnDiagnostics"
                          label="Diagnostics"
                          imageMso="OutlineSlide"
                          size="large"
                          onAction="OnDiagnostics" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;
    }

    public void OnSendSelection(IRibbonControl control)
    {
        // The button captures the current Excel selection and pushes it to the back-end.
        // The actual implementation lives in Commands.SendSelection() so it can be unit-tested.
        Commands.SendSelection();
    }

    public void OnToggleWatcher(IRibbonControl control)
    {
        Commands.ToggleWatcher();
    }

    public void OnDiagnostics(IRibbonControl control)
    {
        Commands.ShowDiagnostics();
    }
}
