using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using BarRaider.SdTools;
using EnvDTE;

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
                var (foregroundHandle, foregroundProcessId) = GetForeground();

                if (IsVisualStudioInstanceFocused(foregroundProcessId))
                {
                    GetDTEInstances((int)foregroundProcessId).FirstOrDefault()?.ExecuteCommand(settings.Command);
                }
                else
                {
                    foreach (var dte in GetDTEInstances())
                    {
                        if (IsFocusingDebugProcess(foregroundProcessId, dte))
                        {
                            dte.ExecuteCommand(settings.Command);

                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.LogMessage(TracingLevel.ERROR, ex.Message);
            }
        }

        private (IntPtr handle, uint processId) GetForeground()
        {
            var foreground = WindowsAPI.GetForegroundWindow();

            WindowsAPI.GetWindowThreadProcessId(foreground, out var pid);

            return (handle: foreground, processId: pid);
        }

        private bool IsVisualStudioInstanceFocused(uint processId)
        {
            return System.Diagnostics.Process.GetProcessById((int)processId)
                .ProcessName.Equals("devenv", StringComparison.InvariantCultureIgnoreCase);
        }

        private static bool IsFocusingDebugProcess(uint processId, DTE dte)
        {
            if (dte.Debugger?.DebuggedProcesses?.Count > 0)
            {
                foreach (EnvDTE.Process debuggedProcess in dte.Debugger.DebuggedProcesses)
                {
                    if (debuggedProcess.ProcessID == processId)
                    {
                        Logger.Instance.LogMessage(TracingLevel.INFO, $"Debugging process found, id: {processId}");

                        return true;
                    }
                }
            }

            return false;
        }

        private IEnumerable<EnvDTE.DTE> GetDTEInstances(int? processId = null)
        {
            var dteInstances = new List<DTE>();
            IBindCtx bindCtx = null;
            IRunningObjectTable runningObjects = null;
            IEnumMoniker monikers = null;

            var foundByProcessId = false;

            try
            {
                Marshal.ThrowExceptionForHR(WindowsAPI.CreateBindCtx(0, out bindCtx));

                bindCtx.GetRunningObjectTable(out runningObjects);

                runningObjects.EnumRunning(out monikers);

                var moniker = new IMoniker[1];
                var fetchedMonikers = IntPtr.Zero;

                while (monikers.Next(1, moniker, fetchedMonikers) == 0)
                {
                    moniker[0].GetDisplayName(bindCtx, null, out var rotName);

                    if (rotName.StartsWith("!VisualStudio.DTE.16.0:") || rotName.StartsWith("!VisualStudio.DTE.15.0:"))
                    {
                        Marshal.ThrowExceptionForHR(runningObjects.GetObject(moniker[0], out var runningObject));

                        if (runningObject is EnvDTE.DTE dte)
                        {
                            Logger.Instance.LogMessage(TracingLevel.INFO, $"ROT Object Found {rotName}");

                            if (processId.HasValue && int.TryParse(rotName.Substring(23), out var rotProcessId) && rotProcessId == processId)
                            {
                                foundByProcessId = true;

                                dteInstances.Clear();
                                dteInstances.Add(dte);

                                break;
                            }

                            dteInstances.Add(dte);
                        }
                    }
                }

                if (processId.HasValue && !foundByProcessId)
                {
                    return Enumerable.Empty<EnvDTE.DTE>();
                }

                return dteInstances.AsEnumerable();
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (runningObjects != null)
                {
                    Marshal.ReleaseComObject(runningObjects);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }
        }
    }
}
