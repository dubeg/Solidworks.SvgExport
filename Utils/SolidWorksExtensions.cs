using SolidWorks.Interop.sldworks;

namespace Dx.Sw.SvgExport.Utils;

public static class SolidWorksExtensions {
    public static void SetPosition(this IView view, double x, double y) => view.Position = new double[] { x, y };

    public static (double X, double Y) GetPositionAsTuple(this IView view) {
        var viewPosition = (double[])view.Position;
        return (
            X: viewPosition[0],
            Y: viewPosition[1]
        );
    }

    public static (double MinX, double MinY, double MaxX, double MaxY) GetOutlineAsTuple(this IView view) {
        var outline = (double[])view.GetOutline();
        var xMin = outline[0];
        var yMin = outline[1];
        var xMax = outline[2];
        var yMax = outline[3];
        return (
            MinX: xMin,
            MinY: yMin,
            MaxX: xMax,
            MaxY: yMax
        );
    }

    public static NxRectangle GetOutlineAsRect(this IView view, bool meterToMM = false) {
        var outline = (double[])view.GetOutline();
        var xMin = outline[0] * (meterToMM ? 1000.0 : 1.0);
        var yMin = outline[1] * (meterToMM ? 1000.0 : 1.0);
        var xMax = outline[2] * (meterToMM ? 1000.0 : 1.0);
        var yMax = outline[3] * (meterToMM ? 1000.0 : 1.0);
        return new NxRectangle() {
            X = xMin,
            Y = yMin,
            Width = xMax - xMin,
            Height = yMax - yMin
        };
    }

    public static (double X, double Y) GetViewCenter(this IView view) {
        var outline = (double[])view.GetOutline();
        var xMin = outline[0];
        var yMin = outline[1];
        var xMax = outline[2];
        var yMax = outline[3];
        return (
            X: (xMax - xMin) / 2.0,
            Y: (yMax - yMin) / 2.0
        );
    }

    public static ISheet GetSheet(this IDrawingDoc swDraw, string sheetName) {
        var sheetNames = (string[])swDraw.GetSheetNames();
        if (sheetNames is null) return null;
        if (sheetNames.Contains(sheetName)) {
            return swDraw.Sheet[sheetName];
        }
        return null;
    }

    public static List<ISheet> GetSheets(this IDrawingDoc swDraw) {
        var results = new List<ISheet>();
        var sheetNames = (string[])swDraw.GetSheetNames();
        if (sheetNames is null) return results;
        foreach (var name in sheetNames) {
            results.Add(swDraw.Sheet[name]);
        }
        return results;
    }

    public static IView GetViewBySheetName(this IDrawingDoc swDraw, string sheetName) {
        var viewsBySheet = (object[])swDraw.GetViews();
        if (viewsBySheet is null) return null;
        foreach (object[] views in viewsBySheet) {
            var view = views.Cast<IView>().FirstOrDefault(x => x.Name == sheetName);
            if (view is not null) {
                return view;
            }
        }
        return null;
    }

    public static List<(IView sheetView, List<IView> InnerViews)> GetViewsBySheet(this IDrawingDoc swDraw) {
        var viewsBySheet = (object[])swDraw.GetViews();
        if (viewsBySheet is null) return null;
        var results = new List<(IView sheetView, List<IView> InnerViews)>();
        foreach (object[] views in viewsBySheet) {
            var sheetViews = views.Cast<IView>().ToList();
            results.Add(
                (sheetViews.First(), sheetViews.Skip(1).ToList())
            );
        }
        return results;
    }

    public static List<IView> GetViewsEx(this IDrawingDoc swDraw) {
        var viewsBySheet = (object[])swDraw.GetViews();
        if (viewsBySheet is null) return null;
        var results = new List<IView>();
        foreach (object[] views in viewsBySheet) {
            results.AddRange(
                views.Cast<IView>()
            );
        }
        return results;
    }

    public static List<IView> GetViewsForSheet(this IDrawingDoc swDraw, ISheet sheet) => GetViewsForSheet(swDraw, sheet.GetName());

    public static List<IView> GetViewsForSheet(this IDrawingDoc swDraw, string sheetName) {
        return swDraw.GetViewsBySheet().FirstOrDefault(x => x.sheetView.GetName2() == sheetName).InnerViews;
    }

    public static IEnumerable<IAnnotation> GetAnnotationsEx(this IView view) {
        var annotations = (object[])view.GetAnnotations();
        if (annotations is not null) {
            return annotations.Cast<IAnnotation>();
        }
        return Enumerable.Empty<IAnnotation>();
    }
}

