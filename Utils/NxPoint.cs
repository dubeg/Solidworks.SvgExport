using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dx.Sw.SvgExport.Utils;

public struct NxPoint {
    public double X;
    public double Y;
    public double Z;
    public NxPoint(double x, double y, double z) {
        X = x;
        Y = y;
        Z = z;
    }
    public double[] ToArray() => [X, Z, Z];
    public static NxPoint FromCoords(double x, double y) => new NxPoint(x, y, 0);
    public static NxPoint FromCoords(double x, double y, double z) => new NxPoint(x, y, z);
    public static NxPoint FromArray(double[] arr) => new NxPoint(arr[0], arr[1], arr[2]);
    public static NxPoint FromSpan(Span<double> span) => new NxPoint(span[0], span[1], span[2]);
}