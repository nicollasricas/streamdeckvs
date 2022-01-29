using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
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
                var (processHandle, processId) = Windows32API.GetForeground();

                if (IsProcessVisualStudio(processId))
                {
                    ExecuteCommand(DTEAPI.GetDTE(processId).FirstOrDefault(), settings);
                }
                else
                {
                    var dte = DTEAPI.GetDTE().FirstOrDefault(m => IsProcessAttachedToDebug(processId, m));

                    if (dte is null)
                    {
                        var processCommandLine = GetProcessCommandLine(processId);

                        if (IsLinkedByPipe(processCommandLine))
                        {
                            ExecuteCommand(DTEAPI.GetDTE(GetVisualStudioPIDFromPipeLink(processCommandLine))
                                .FirstOrDefault(), settings);
                        }
                    }
                    else
                    {
                        ExecuteCommand(dte, settings);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.LogMessage(TracingLevel.ERROR, ex.Message);
            }
        }

        private int GetVisualStudioPIDFromPipeLink(string link)
        {
            var indexLength = @"\\.\pipe\Microsoft-VisualStudio-Debug-Console-".Length;

            var start = link.IndexOf(@"\\.\pipe\Microsoft-VisualStudio-Debug-Console-");

            if (start >= 0)
            {
                var end = link.IndexOf(" ", start);

                if (end >= 0)
                {
                    if (int.TryParse(link.Substring(start + indexLength, end - start - indexLength), out var pid))
                    {
                        return pid;
                    }
                }
            }

            return 0;
        }

        private bool IsLinkedByPipe(string arguments) => arguments.Contains(@"\\.\pipe\Microsoft-VisualStudio-Debug-Console-");

        private void ExecuteCommand(DTE dte, ExecuteCommandSettings settings) => dte?.ExecuteCommand(settings.Command,settings.CommandArgs);

        private bool IsProcessVisualStudio(int processId)
        {
            return System.Diagnostics.Process.GetProcessById(processId)
                .ProcessName.Equals("devenv", StringComparison.InvariantCultureIgnoreCase);
        }

        private static bool IsProcessAttachedToDebug(int processId, DTE dte)
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

        private string GetProcessCommandLine(int processId)
        {
            try
            {
                var wmi = new ManagementObjectSearcher("root\\CIMV2", $"SELECT * FROM Win32_Process where ProcessId = {processId}");

                var process = Enumerable.FirstOrDefault(wmi.Get() as IEnumerable<ManagementObject>);

                return (string)process?["CommandLine"];
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
