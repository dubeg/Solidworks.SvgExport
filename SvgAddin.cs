using CADBooster.SolidDna;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Dx.Sw.SvgExport;

[Guid("652bcb12-9d20-4010-b47e-026f1ac07ce3"), ComVisible(true)]
public class SvgAddin : SolidAddIn {
    /// <summary>
    /// Specific application start-up code
    /// </summary>
    public override void ApplicationStartup() {
    }

    /// <summary>
    /// Use this to do early initialization and any configuration of the 
    /// PlugInIntegration class properties.
    /// </summary>
    public override void PreConnectToSolidWorks() {
        Logger.AddFileLogger<SvgAddin>("logs.log", new FileLoggerConfiguration() {
            LogLevel = LogLevel.Trace
        });
    }

    /// <summary>
    /// Steps to take before any plug-in loads
    /// </summary>
    public override void PreLoadPlugIns() {}
}
