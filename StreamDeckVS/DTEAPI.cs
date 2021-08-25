using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using BarRaider.SdTools;
using EnvDTE;

namespace StreamDeckVS
{
    public static class DTEAPI
    {
        public static IEnumerable<EnvDTE.DTE> GetDTE(int? processId = null)
        {
            var dteInstances = new List<DTE>();
            IBindCtx bindCtx = null;
            IRunningObjectTable runningObjects = null;
            IEnumMoniker monikers = null;

            var foundByProcessId = false;

            try
            {
                Marshal.ThrowExceptionForHR(Windows32API.CreateBindCtx(0, out bindCtx));

                bindCtx.GetRunningObjectTable(out runningObjects);

                runningObjects.EnumRunning(out monikers);

                var moniker = new IMoniker[1];
                var fetchedMonikers = IntPtr.Zero;

                while (monikers.Next(1, moniker, fetchedMonikers) == 0)
                {
                    moniker[0].GetDisplayName(bindCtx, null, out var rotName);

                    if (rotName.StartsWith("!VisualStudio.DTE.17.0:") || rotName.StartsWith("!VisualStudio.DTE.16.0:") || rotName.StartsWith("!VisualStudio.DTE.15.0:"))
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
