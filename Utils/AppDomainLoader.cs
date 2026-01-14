using System;
using System.IO;
using System.Reflection;
using TypeNameResolver;

namespace Dx.Sw.SvgExport.Utils;

public static class AppDomainLoader {
    public static void Init() {
        AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AddInAssemblyResolveEventHandler);
    }

    /// <summary>
    /// Required for loading the add-in.
    /// Refer to the issue here:
    /// https://stackoverflow.com/questions/72978989/
    /// </summary>
    private static Assembly AddInAssemblyResolveEventHandler(object sender, ResolveEventArgs args) {
        var currAssemblyPath = Assembly.GetExecutingAssembly().Location;
        var currAssemblyDir = Path.GetDirectoryName(currAssemblyPath);
        var assemblyPath = Assembly.GetCallingAssembly().Location;
        var assemblyDir = Path.GetDirectoryName(assemblyPath);
        if (currAssemblyDir == assemblyDir) {
            var lib = TypeNameParser.Parse(args.Name);
            var libName = lib.Scope.FullName;
            var libPath = Path.Combine(currAssemblyDir, $"{libName}.dll");
            if (File.Exists(libPath))
                return Assembly.LoadFile(libPath);
        }
        return null;
    }
}
