using System.Collections.Generic;

namespace Dx.Sw.SvgExport.Utils;

/// <summary>
/// Holds parsed data from a single polyline record returned by GetPolylines6.
/// </summary>
public struct NxPolylineRecord {
    public NxPolylineType Type;
    /// <summary>For arcs: [centerX, centerY, centerZ, startX, startY, startZ, endX, endY, endZ, normalX, normalY, normalZ]</summary>
    public double[] GeomData;
    public int LineColor;
    public int LineStyle;
    public int LineFont;
    public double LineWeight;
    public int LayerID;
    public int LayerOverride;
    /// <summary>Tessellated points as (x, y, z) triplets in meters.</summary>
    public List<NxPoint> Points;
}