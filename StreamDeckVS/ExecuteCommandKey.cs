using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using BarRaider.SdTools;

namespace StreamDeckVS
{
    [PluginActionId("com.nicollasr.streamdeckvs.executecommand")]
    public class ExecuteCommandKey : Key<ExecuteCommandSettings>
    {
        public ExecuteCommandKey(SDConnection connection, InitialPayload payload) : base(connection, payload)
        {
        }

        public override void KeyPressed(KeyPayload payload)
        {
            base.KeyPressed(payload);

            try
            {
                GetDTE()?.ExecuteCommand(settings.Command);
            }
            catch (Exception ex)
            {
                Logger.Instance.LogMessage(TracingLevel.ERROR, ex.Message);
            }
        }

        private EnvDTE.DTE GetDTE()
        {
            var window = WindowsAPI.GetForegroundWindow();

            WindowsAPI.GetWindowThreadProcessId(window, out var processPid);

            if (processPid < 0)
            {
                return null;
            }

            Logger.Instance.LogMessage(TracingLevel.INFO, $"Foreground Process PID: {processPid}");

            Marshal.ThrowExceptionForHR(WindowsAPI.CreateBindCtx(0, out var bindContext));

            bindContext.GetRunningObjectTable(out var runningTable);

            runningTable.EnumRunning(out var monikers);

            var moniker = new IMoniker[1];
            var fetchedMonikers = IntPtr.Zero;

            while (monikers.Next(1, moniker, fetchedMonikers) == 0)
            {
                moniker[0].GetDisplayName(bindContext, null, out var runningObjectName);

                if (runningObjectName == $"!VisualStudio.DTE.16.0:{processPid}" || runningObjectName == $"!VisualStudio.DTE.15.0:{processPid}")
                {
                    Marshal.ThrowExceptionForHR(runningTable.GetObject(moniker[0], out var runningObject));

                    if (monikers != null)
                    {
                        Marshal.ReleaseComObject(monikers);
                    }

                    if (runningTable != null)
                    {
                        Marshal.ReleaseComObject(runningTable);
                    }

                    if (bindContext != null)
                    {
                        Marshal.ReleaseComObject(bindContext);
                    }

                    Logger.Instance.LogMessage(TracingLevel.INFO, $"ROT Object found: {runningObjectName}");

                    return runningObject as EnvDTE.DTE;
                }
            }

            return null;
        }
    }
}
