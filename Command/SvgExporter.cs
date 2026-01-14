using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Dx.Sw.SvgExport.Utils;
using MathNet.Numerics.LinearAlgebra;

namespace Dx.Sw.SvgExport.Command;

public class SvgExporter(ISldWorks _app) {
    private bool _drawDetailCircle => true;
    private bool _drawDetailCircleArrows => false;
    private bool _drawDetailCircleLabel => false;
    private bool _drawDetailViewLabel => false;
    private bool _drawDetailViewConnectingLine => true;

    public List<string> Export(
        string filePath, 
        bool fitToContent = false, 
        bool includeBomMetadata = false, 
        bool exportIndividualViews = false, 
        string specificViewName = null, 
        bool excludeViewsOutOfBounds = false
    ) {
        double MeterToMM(double value) => value * 1000.0;
        var warnings = new List<string>();
        // --
        var model = _app.IActiveDoc2;
        if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING) {
            throw new InvalidOperationException("Active document is not a drawing.");
        }
        var drawing = (DrawingDoc)model;
        var sheet = drawing.IGetCurrentSheet();
        var sheetView = drawing.GetViewBySheetName(sheet.GetName());
        var sheetScale2 = sheetView.ScaleDecimal;
        var props = (double[])sheet.GetProperties2(); // [ paperSize, templateIn, scale1, scale2, firstAngle, width, height, ? ]
        var sheetScale = props[2] / props[3];
        var sheetWidth = MeterToMM(props[5]);
        var sheetHeight = MeterToMM(props[6]);
        var globalWriter = new SvgWriter(sheetWidth, sheetHeight);
        // Y-flip transform matrix: flips Y-axis and translates to keep coordinates positive
        // | 1   0   0  |   | x |   | x            |
        // | 0  -1   H  | × | y | = | H - y        |
        // | 0   0   1  |   | 1 |   | 1            |
        var yFlipTransform = Matrix<double>.Build.DenseOfArray(new double[,] {
            { 1,  0,  0 },
            { 0, -1,  sheetHeight },
            { 0,  0,  1 }
        });
        (double x, double y) ApplyTransformY(double x, double y) {
            var pt = Vector<double>.Build.Dense(new[] { x, y, 1.0 });
            var result = yFlipTransform * pt;
            return (result[0], result[1]);
        }
        var bomData = includeBomMetadata ? ParseBomTable(drawing) : new Dictionary<string, BomItem>();
        var views = drawing.GetViewsForSheet(sheet);
        
        // Filter views if a specific view name is provided
        if (!string.IsNullOrEmpty(specificViewName)) {
            views = views.Where(v => v.GetName2() == specificViewName).ToList();
            if (views.Count == 0) {
                warnings.Add($"Warning: Specified view '{specificViewName}' not found in the current sheet.");
                return warnings;
            }
        }
        
        var viewCount = views.Count;
        var isSingleView = viewCount == 1;
        var viewIndex = 1;
        foreach (var view in views) {
            var visible = view.GetVisible();
            if (!visible) continue;
            // Check if view is entirely out of bounds (no intersection with sheet) and should be excluded
            if (excludeViewsOutOfBounds) {
                var viewRect = view.GetOutlineAsRect(meterToMM: true);
                var isEntirelyOutOfBounds = 
                    (viewRect.X + viewRect.Width) <= 0   // entirely to the left of sheet
                    || viewRect.X >= sheetWidth          // entirely to the right of sheet
                    || (viewRect.Y + viewRect.Height) <= 0  // entirely below sheet
                    || viewRect.Y >= sheetHeight;        // entirely above sheet
                if (isEntirelyOutOfBounds) {
                    var viewName = view.GetName2();
                    warnings.Add($"Excluded view '{viewName}' because it is entirely out of sheet bounds.");
                    continue;
                }
            }
            if (isSingleView) {
                RenderView(view, globalWriter, [globalWriter], bomData, includeBomMetadata, sheetHeight, ApplyTransformY, warnings);
            } 
            else {
                // Multiple views: create individual view writer
                var viewName = view.GetName2();
                var singleViewWriter = new SvgWriter(sheetWidth, sheetHeight) { DocumentName = viewName };
                RenderView(view, globalWriter, [globalWriter, singleViewWriter], bomData, includeBomMetadata, sheetHeight, ApplyTransformY, warnings);
                if (fitToContent) {
                    singleViewWriter.ApplyFitToContent(padding: 5.0);
                }
                if (exportIndividualViews) {
                    var directory = System.IO.Path.GetDirectoryName(filePath);
                    var fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    var viewFilePath = System.IO.Path.Combine(directory, $"{fileNameWithoutExt}_v{viewIndex}.svg");
                    singleViewWriter.Save(viewFilePath);
                    viewIndex++;
                }
            }
        }
        if (fitToContent) {
            globalWriter.ApplyFitToContent(padding: 5.0);
        }
        globalWriter.Save(filePath);
        return warnings;
    }

    /// <summary>
    /// Renders a single drawing view to the specified SVG writer.
    /// </summary>
    private void RenderView(
        IView view, 
        SvgWriter mainSvgWriter,
        SvgWriter[] writers, 
        Dictionary<string, BomItem> bomData,
        bool includeBomMetadata,
        double sheetHeight,
        Func<double, double, (double, double)> applyTransformY,
        List<string> warnings
    ) {
        double MeterToMM(double value) => value * 1000.0;
        var viewName = view.GetName2(); // For debugging.
        var viewType = (swDrawingViewTypes_e)view.Type;
        // Get the view's origin position on the sheet (where sketch origin projects to)
        var (viewOriginX, viewOriginY) = view.GetPositionAsTuple();
        viewOriginX = MeterToMM(viewOriginX);
        viewOriginY = MeterToMM(viewOriginY);
        var viewRect = view.GetOutlineAsRect(meterToMM: true);
        var viewLeft = viewOriginX - viewRect.Width / 2;
        var viewBottom = viewOriginY - viewRect.Height / 2;
        // Transform point from view/sketch space (meters) to SVG space (mm, Y-flipped)
        (double svgX, double svgY) ViewToSvg(double x, double y) {
            var mmX = MeterToMM(x) * view.ScaleDecimal + viewOriginX;
            var mmY = MeterToMM(y) * view.ScaleDecimal + viewOriginY;
            (mmX, mmY) = applyTransformY(mmX, mmY);
            return (mmX, mmY);
        }
        (double svgX, double svgY) SheetToSvg(double x, double y) {
            var mmX = MeterToMM(x);
            var mmY = MeterToMM(y);
            (mmX, mmY) = applyTransformY(mmX, mmY);
            return (mmX, mmY);
        }
        var debugColor = "#FF0000"; // Red
        // -----------------
        // View with detail circles
        // -----------------
        // Transform coordinates from sheet space (meters) to SVG space
        // Detail circle coordinates are in sheet space, like annotations.
        (double svgX, double svgY) DetailCircleToSvg(double x, double y) => SheetToSvg(x, y);
        // TODO: fetch from API
        var detailColor = "#000000"; // Default black
        // TODO: fetch from API
        var detailStrokeWidth = 0.25; // Default stroke width
        if (_drawDetailCircle) {
            // There will be a main view (vue de mise en plan) with multiple detail circles linked to other Detail Views.
            // In that main view, GetDetailCircleInfo2 will return something.
            // In the other Detail View(s), it won't return anything.
            var detailCircleCount = view.GetDetailCircleCount2(out var arrSize);
            if (detailCircleCount > 0) {
                var detailCircleInfos = (double[])view.GetDetailCircleInfo2();
                var detailCircleLabels = (string[])view.GetDetailCircleStrings();
                // Index 0: number of detail circles
                // Index 1: data
                var index = 1; 
                for (var i = 0; i < detailCircleCount; i++) {
                    // -------------------------
                    // Circle
                    // -------------------------
                    var layer = (int)detailCircleInfos[index + 0];
                    var centerX = detailCircleInfos[index + 1];
                    var centerY = detailCircleInfos[index + 2];
                    var centerZ = detailCircleInfos[index + 3];
                    var startX = detailCircleInfos[index + 4];
                    var startY = detailCircleInfos[index + 5];
                    var startZ = detailCircleInfos[index + 6];
                    var endX = detailCircleInfos[index + 7];
                    var endY = detailCircleInfos[index + 8];
                    var endZ = detailCircleInfos[index + 9];
                    var lineType = (swLineTypes_e)(int)detailCircleInfos[index + 10];
                    var textX = detailCircleInfos[index + 11];
                    var textY = detailCircleInfos[index + 12];
                    var textZ = detailCircleInfos[index + 13];
                    var textHeight = detailCircleInfos[index + 14]; // in meters
                    var numArrows = (int)detailCircleInfos[index + 15];
                    index += 16;
                    var radius = Math.Sqrt(Math.Pow(startX - centerX, 2) + Math.Pow(startY - centerY, 2));
                    var radiusMM = MeterToMM(radius);
                    var (svgCenterX, svgCenterY) = DetailCircleToSvg(centerX, centerY);
                    var (svgTextX, svgTextY) = DetailCircleToSvg(textX, textY);
                    mainSvgWriter.StartGroup(className: "annotation annotation-detail-circle");
                    mainSvgWriter.AddCircle(svgCenterX, svgCenterY, radiusMM, detailColor, detailStrokeWidth);
                    mainSvgWriter.UpdateBounds(svgCenterX - radiusMM, svgCenterY - radiusMM);
                    mainSvgWriter.UpdateBounds(svgCenterX + radiusMM, svgCenterY + radiusMM);
                    // -------------------------
                    // Arrows on broken circle
                    // -------------------------
                    if (_drawDetailCircleArrows && true) {
                        for (var j = 0; j < numArrows; j++) {
                            var arrowTipX = detailCircleInfos[index + 0];
                            var arrowTipY = detailCircleInfos[index + 1];
                            var arrowTipZ = detailCircleInfos[index + 2];
                            var arrowCompX = detailCircleInfos[index + 3];
                            var arrowCompY = detailCircleInfos[index + 4];
                            var arrowCompZ = detailCircleInfos[index + 5];
                            var arrowWidth = detailCircleInfos[index + 6]; // in meters
                            var arrowHeight = detailCircleInfos[index + 7]; // in meters
                            var arrowStyle = (swArrowStyle_e)(int)detailCircleInfos[index + 8];
                            index += 9;
                            // Move to next arrow
                            // Convert arrow points to SVG coordinates
                            var (svgTipX, svgTipY) = DetailCircleToSvg(arrowTipX, arrowTipY);
                            var (svgCompX, svgCompY) = DetailCircleToSvg(arrowCompX, arrowCompY);
                            var arrowWidthMM = MeterToMM(arrowWidth);
                            var arrowHeightMM = MeterToMM(arrowHeight);
                            // Calculate direction vector from comp to tip
                            var dirX = svgTipX - svgCompX;
                            var dirY = svgTipY - svgCompY;
                            var dirLength = Math.Sqrt(dirX * dirX + dirY * dirY);
                            if (dirLength > 0) {
                                // Normalize direction
                                dirX /= dirLength;
                                dirY /= dirLength;
                                // Perpendicular vector for arrow width
                                var perpX = -dirY;
                                var perpY = dirX;
                                // Calculate arrow base point (back from tip by arrow height)
                                var baseX = svgTipX - dirX * arrowHeightMM;
                                var baseY = svgTipY - dirY * arrowHeightMM;
                                // Calculate the two base corners of the arrow (half width on each side)
                                var halfWidth = arrowWidthMM / 2;
                                var corner1X = baseX + perpX * halfWidth;
                                var corner1Y = baseY + perpY * halfWidth;
                                var corner2X = baseX - perpX * halfWidth;
                                var corner2Y = baseY - perpY * halfWidth;
                                // Draw filled triangular arrowhead
                                var arrowPath = $"M {Format(svgTipX)} {Format(svgTipY)} L {Format(corner1X)} {Format(corner1Y)} L {Format(corner2X)} {Format(corner2Y)} Z";
                                mainSvgWriter.AddPath(arrowPath, debugColor, detailStrokeWidth, debugColor);
                            }
                        }
                    }
                    else {
                        for (var j = 0; j < numArrows; j++) {
                            index += 9;
                        }
                    }
                    mainSvgWriter.EndGroup(); // End annotation-detail-circle
                }
            }
            var detailCirclesArr = view.GetDetailCircles();
            if (detailCirclesArr is not null) {
                var detailCircles = ((object[])detailCirclesArr).Cast<IDetailCircle>();
                foreach (var detailCircle in detailCircles) {
                    var detailCircleName = detailCircle.GetName();
                    var typeOfCircle = (swDetCircleShowType_e)detailCircle.GetDisplay();
                    var detailCircleStyle = (swDetViewStyle_e)detailCircle.GetStyle();
                    // ---------------------------
                    // Label
                    // ---------------------------
                    if (_drawDetailCircleLabel && true) {
                        var detailCircleLabel = detailCircle.GetLabel();
                        detailCircle.GetLabelPosition(out var labelX, out var labelY);
                        var (lblx, lbly) = DetailCircleToSvg(labelX, labelY);
                        var textFormat = detailCircle.GetTextFormat();
                        var textHeightMM = textFormat.CharHeight * 1000.0; // Meters to MM
                        mainSvgWriter.StartGroup(className: "annotation annotation-detail-circle-label");
                        mainSvgWriter.AddText(lblx, lbly, detailCircleLabel, "Arial", textHeightMM, debugColor, 0, "middle");
                        mainSvgWriter.EndGroup();
                    }
                    // ---------------------------
                    // Connecting line(s) between detail circle & detail view.
                    // ---------------------------
                    if (_drawDetailViewConnectingLine) {
                        if (detailCircleStyle == swDetViewStyle_e.swDetViewCONNECTED) {
                            var connectingLine = (double[])detailCircle.GetConnectingLine();
                            var pointA = NxPoint.FromSpan(connectingLine.AsSpan(0, 3));
                            var pointB = NxPoint.FromSpan(connectingLine.AsSpan(3, 3));
                            var (ax, ay) = DetailCircleToSvg(pointA.X, pointA.Y);
                            var (bx, by) = DetailCircleToSvg(pointB.X, pointB.Y);
                            mainSvgWriter.StartGroup(className: "annotation annotation-connecting-line");
                            mainSvgWriter.AddLine(ax, ay, bx, by, detailColor, detailStrokeWidth);
                            mainSvgWriter.EndGroup();
                        }
                    }
                }
            }
            // ------------
            // Section lines (I don't know what they are yet)
            // ------------
            if (false) {
                var sectionLineCount = view.GetSectionLineCount2(out var sectionLineInfoArrSize);
                if (sectionLineCount > 0) {
                    var sectionLineInfoArr = view.GetSectionLineInfo2();
                    var sectionLineStringsArr = view.GetSectionLineStrings();
                    var sectionLineStrings = sectionLineStringsArr is not null ? (string[])sectionLineStringsArr : Array.Empty<string>();
                    if (sectionLineInfoArr is not null) {
                        var sectionLineInfo = (double[])sectionLineInfoArr;
                        // Array structure: [numSectionLines, layer, [numSegments, [lineType, startPt[3], endPt[3]], 
                        //   arrowStart1[3], arrowEnd1[3], arrowWidth1, arrowHeight1, arrowStyle1,
                        //   arrowStart2[3], arrowEnd2[3], arrowWidth2, arrowHeight2, arrowStyle2,
                        //   textPt1[3], textPt2[3], textHeight ] ]
                        var debugStrokeWidth = 0.5;
                        var index = 0;
                        var numSectionLines = (int)sectionLineInfo[index++];
                        var layer = (int)sectionLineInfo[index++];
                        for (var sl = 0; sl < numSectionLines; sl++) {
                            var numSegments = (int)sectionLineInfo[index++];
                            // Draw line segments
                            for (var seg = 0; seg < numSegments; seg++) {
                                var lineType = (swLineTypes_e)(int)sectionLineInfo[index++];
                                // startPt[3]
                                var startX = sectionLineInfo[index++];
                                var startY = sectionLineInfo[index++];
                                var startZ = sectionLineInfo[index++];
                                // endPt[3]
                                var endX = sectionLineInfo[index++];
                                var endY = sectionLineInfo[index++];
                                var endZ = sectionLineInfo[index++];
                                var (sx, sy) = SheetToSvg(startX, startY);
                                var (ex, ey) = SheetToSvg(endX, endY);
                                mainSvgWriter.AddLine(sx, sy, ex, ey, debugColor, debugStrokeWidth);
                                mainSvgWriter.UpdateBounds(sx, sy);
                                mainSvgWriter.UpdateBounds(ex, ey);
                            }
                            // Arrow 1 data
                            var arrowStart1X = sectionLineInfo[index++];
                            var arrowStart1Y = sectionLineInfo[index++];
                            var arrowStart1Z = sectionLineInfo[index++];
                            var arrowEnd1X = sectionLineInfo[index++];
                            var arrowEnd1Y = sectionLineInfo[index++];
                            var arrowEnd1Z = sectionLineInfo[index++];
                            var arrowWidth1 = sectionLineInfo[index++];
                            var arrowHeight1 = sectionLineInfo[index++];
                            var arrowStyle1 = (swArrowStyle_e)(int)sectionLineInfo[index++];
                            // Arrow 2 data
                            var arrowStart2X = sectionLineInfo[index++];
                            var arrowStart2Y = sectionLineInfo[index++];
                            var arrowStart2Z = sectionLineInfo[index++];
                            var arrowEnd2X = sectionLineInfo[index++];
                            var arrowEnd2Y = sectionLineInfo[index++];
                            var arrowEnd2Z = sectionLineInfo[index++];
                            var arrowWidth2 = sectionLineInfo[index++];
                            var arrowHeight2 = sectionLineInfo[index++];
                            var arrowStyle2 = (swArrowStyle_e)(int)sectionLineInfo[index++];
                            // Text positions
                            var textPt1X = sectionLineInfo[index++];
                            var textPt1Y = sectionLineInfo[index++];
                            var textPt1Z = sectionLineInfo[index++];
                            var textPt2X = sectionLineInfo[index++];
                            var textPt2Y = sectionLineInfo[index++];
                            var textPt2Z = sectionLineInfo[index++];
                            var textHeight = sectionLineInfo[index++];
                            // Draw arrow 1 (line from start to end point of arrow)
                            var (a1sx, a1sy) = SheetToSvg(arrowStart1X, arrowStart1Y);
                            var (a1ex, a1ey) = SheetToSvg(arrowEnd1X, arrowEnd1Y);
                            mainSvgWriter.AddLine(a1sx, a1sy, a1ex, a1ey, debugColor, debugStrokeWidth);
                            mainSvgWriter.UpdateBounds(a1sx, a1sy);
                            mainSvgWriter.UpdateBounds(a1ex, a1ey);
                            // Draw arrow 2 (line from start to end point of arrow)
                            var (a2sx, a2sy) = SheetToSvg(arrowStart2X, arrowStart2Y);
                            var (a2ex, a2ey) = SheetToSvg(arrowEnd2X, arrowEnd2Y);
                            mainSvgWriter.AddLine(a2sx, a2sy, a2ex, a2ey, debugColor, debugStrokeWidth);
                            mainSvgWriter.UpdateBounds(a2sx, a2sy);
                            mainSvgWriter.UpdateBounds(a2ex, a2ey);
                            // Draw text labels at both ends
                            var textHeightMM = MeterToMM(textHeight);
                            var fontSize = textHeightMM > 0 ? textHeightMM : 3.5;
                            // Each section line has two labels (one at each arrow)
                            var labelIndex = sl * 2;
                            if (labelIndex < sectionLineStrings.Length) {
                                var label1 = sectionLineStrings[labelIndex];
                                if (!string.IsNullOrEmpty(label1)) {
                                    var (t1x, t1y) = SheetToSvg(textPt1X, textPt1Y);
                                    mainSvgWriter.AddText(t1x, t1y, label1, "Arial", fontSize, debugColor, 0, "middle");
                                    mainSvgWriter.UpdateBounds(t1x, t1y);
                                }
                            }
                            if (labelIndex + 1 < sectionLineStrings.Length) {
                                var label2 = sectionLineStrings[labelIndex + 1];
                                if (!string.IsNullOrEmpty(label2)) {
                                    var (t2x, t2y) = SheetToSvg(textPt2X, textPt2Y);
                                    mainSvgWriter.AddText(t2x, t2y, label2, "Arial", fontSize, debugColor, 0, "middle");
                                    mainSvgWriter.UpdateBounds(t2x, t2y);
                                }
                            }
                        }
                    }
                }
            }
            // ...
        }
        // -------------
        // Detail View
        // -------------
        if (viewType == swDrawingViewTypes_e.swDrawingDetailView) {
            // ---------------------------
            // Label (view name | label)
            // ---------------------------
            if (_drawDetailViewLabel) {
                var detailAnnotations = view.GetAnnotationsEx();
                foreach (var ann in detailAnnotations) {
                    var annotationVisible = (swAnnotationVisibilityState_e)ann.Visible;
                    if (annotationVisible != swAnnotationVisibilityState_e.swAnnotationVisible) continue;
                    var annType = (swAnnotationType_e)ann.GetType();
                    if (annType != swAnnotationType_e.swNote) continue;
                    var note = (INote)ann.GetSpecificAnnotation();
                    if (note == null || note.IsBomBalloon()) continue;
                    var noteText = note.GetText();
                    var noteName = note.GetName();
                    var propertyLinkedText = note.PropertyLinkedText;
                    var isDetailViewLabel =
                        !string.IsNullOrWhiteSpace(propertyLinkedText)
                        && (
                            propertyLinkedText.Contains("VLNAME")
                            || propertyLinkedText.Contains("VLLABEL")
                        );
                    if (!isDetailViewLabel) continue;
                    
                    // Get default text format for font family and styling
                    var textFormat = (ITextFormat)ann.GetTextFormat(0);
                    var typeFaceName = textFormat.TypeFaceName;
                    var bold = textFormat.Bold;
                    var italic = textFormat.Italic;
                    
                    mainSvgWriter.StartGroup(className: "annotation annotation-detail-view-label");
                    var textCount = note.GetTextCount();
                    if (textCount > 1) {
                        // Multi-line text: render each line at its exact position
                        for (var i = 1; i <= textCount; i++) {
                            var lineText = note.GetTextAtIndex(i);
                            if (string.IsNullOrEmpty(lineText)) continue;
                            var linePosition = (double[])note.GetTextPositionAtIndex(i); // relative to sheet origin
                            if (linePosition == null) continue;
                            var (lineSvgX, lineSvgY) = SheetToSvg(linePosition[0], linePosition[1]);
                            var lineHeight = note.GetTextHeightAtIndex(i) * 1000.0; // Meters to MM
                            var refPosition = (swTextPosition_e)note.GetTextRefPositionAtIndex(i);
                            // var justification = (swTextJustification_e)note.GetTextJustificationAtIndex(i);
                            var (dominantBaseline, textAnchor) = MapRefPositionToBaseline(refPosition);
                            mainSvgWriter.AddText(lineSvgX, lineSvgY, lineText, typeFaceName, lineHeight, debugColor, 0, textAnchor, dominantBaseline, bold, italic);
                        }
                    } 
                    else {
                        // Single-line text
                        if (string.IsNullOrEmpty(noteText)) {
                            mainSvgWriter.EndGroup();
                            continue;
                        }
                        var pos = (double[])ann.GetPosition();
                        if (pos == null) {
                            mainSvgWriter.EndGroup();
                            continue;
                        }
                        var (tx, ty) = SheetToSvg(pos[0], pos[1]);
                        
                        var textJustification = (swTextJustification_e)note.GetTextJustification();
                        var charHeightMM = textFormat.CharHeight * 1000.0; // Meters to MM
                        
                        // For single-line, use note-level justification
                        var textAnchor = MapJustificationToTextAnchor(textJustification);
                        
                        mainSvgWriter.AddText(tx, ty, noteText, typeFaceName, charHeightMM, debugColor, 0, textAnchor, "hanging", bold, italic);
                    }
                    mainSvgWriter.EndGroup();
                }
            }
        }
        // -----------------
        // Polylines
        // -----------------
        // Warning: never use GetPolylines7, as it seems to go through each sub model & it is slow as hell.
        // It may even become unresponsive on a drawing with a large assembly.
        // It seems to go through each sub model of the assembly (if any) & dispatch calls to the UI as I could see sketch segments
        // getting selected one by one while this method was called.
        // Meanwhile, GetPolylines6 returns instantly in all cases.
        var model = (IDrawingDoc)_app.IActiveDoc2;
        var edges = (object[])view.GetPolylines6((int)swCrossHatchFilter_e.swCrossHatchExclude, out var oPolylinesData);
        var polylinesData = (double[])oPolylinesData;
        if (polylinesData == null || polylinesData.Length == 0) return;
        var records = ParsePolylines(polylinesData);
        foreach (var record in records) {
            if (record.Points == null || record.Points.Count < 2) continue;
            var color = ColorRefToHex(record.LineColor);
            var strokeWidth = MapLineWeight((int)record.LineWeight);
            double[] segments = null;
            if (record.LineStyle == -1) {
                // TODO: implement.
                // User has manually edited the line.
                // var lineFontName = model.GetLineFontName2(record.LineFont); // Hidden, Visible, etc.
                // var lineFontInfo = model.GetLineFontInfo2(record.LineFont); // Array containing the line font information 
            }
            else {
                // LineStyle Scale Factor ???
                // 5:1 Ratio according to Claude. Why?
                const double lineStyleScaleFactor = 4.5; 
                // FontInfo: array of double
                // [0]: line weight
                // [1]: segment count
                // [2..segCount]: length of each segment
                // * If length of a segment is negative, this indicates space. Otherwise, it's a line segment.
                var fontName = model.GetLineFontName(record.LineStyle); // Visible, TanVisible, Explodelines, ...
                var fontInfo = (double[])model.GetLineFontInfo(record.LineStyle);
                var lineWeight = fontInfo[0];
                var segmentCount = (int)fontInfo[1];
                if (segmentCount > 1) {
                    segments = fontInfo.AsSpan(2, segmentCount).ToArray();
                    for (var i = 0; i < segments.Length; i++) {
                        segments[i] = Math.Abs(segments[i]) * lineStyleScaleFactor;
                    }
                }
            }
            // ---------------------
            // Polyline
            // ---------------------
            var points = record.Points;
            if (points == null || points.Count < 2) return;
            var pathBuilder = new System.Text.StringBuilder();
            for (var i = 0; i < points.Count; i++) {
                var isFirst = i == 0;
                var pt = points[i];
                var (svgX, svgY) = ViewToSvg(pt.X, pt.Y);
                pathBuilder.Append($" {(isFirst ? "M" : "L")} {Format(svgX)} {Format(svgY)}");
                foreach(var writer in writers) writer.UpdateBounds(svgX, svgY);
            }
            foreach (var writer in writers) writer.AddPath(pathBuilder.ToString(), color, strokeWidth, strokeDashArray: segments);
        }
        // ---------------------------------------------------------
        // Break lines
        // ---------------------------------------------------------
        (double svgX, double svgY) BreakLinePointToSvg(double x, double y) {
            var mmX = MeterToMM(x);
            var mmY = MeterToMM(y);
            (mmX, mmY) = applyTransformY(mmX, mmY);
            return (mmX, mmY);
        }
        if (view.IsBroken()) {
            //var breakLinesArr = (object[])view.GetBreakLines();
            //var breakLines = breakLinesArr.Cast<IBreakLine>();
            //foreach (var breakLine in breakLines) {
            //    var breakLineOrientation = (swBreakLineOrientation_e)breakLine.Orientation;
            //    var breakLinePosition1 = breakLine.GetPosition(0);
            //    var breakLinePosition2 = breakLine.GetPosition(1);
            //    var breakLineGap = view.BreakLineGap;
            //    var hasXAxisBreak = false;
            //    double? breakXAxis_x1 = null;
            //    double? breakXAxis_x2 = null;
            //    var hasYAxisBreak = false;
            //    double? breakYAxis_y1 = null;
            //    double? breakYAxis_y2 = null;
            //    switch (breakLineOrientation) {
            //        case swBreakLineOrientation_e.swBreakLineHorizontal:
            //            breakYAxis_y1 = breakLinePosition1;
            //            breakYAxis_y2 = breakLinePosition2;
            //            hasYAxisBreak = true;
            //            break;
            //        case swBreakLineOrientation_e.swBreakLineVertical:
            //            breakXAxis_x1 = breakLinePosition1;
            //            breakXAxis_x2 = breakLinePosition2;
            //            hasXAxisBreak = true;
            //            break;
            //    }
            //    // Convert break line positions from view-rel space to sheet space.
            //    // Break line positions from GetPosition() are in view-relative space (meters),
            //    // but annotation positions from ann.GetPosition() are in sheet space (meters).
            //    // ie. their origins are different.
            //    var viewPosMeters = (double[])view.Position;
            //    var viewOriginXMeters = viewPosMeters[0];
            //    var viewOriginYMeters = viewPosMeters[1];
            //    var viewScale = view.ScaleDecimal;
            //    if (hasXAxisBreak) {
            //        breakXAxis_x1 = breakXAxis_x1 + viewOriginXMeters;
            //        breakXAxis_x2 = breakXAxis_x2 + viewOriginXMeters;
            //    }
            //    if (hasYAxisBreak) {
            //        breakYAxis_y1 = breakYAxis_y1 + viewOriginYMeters;
            //        breakYAxis_y2 = breakYAxis_y2 + viewOriginYMeters;
            //    }
            var breakCount = view.GetBreakLineCount2(out var breakLineArraySize);
            if (breakCount >= 1) {
                var breakData = (double[])view.GetBreakLineInfo2();
                if (breakData != null && breakData.Length >= 10) {
                    foreach (var writer in writers) writer.StartGroup(className: "annotation annotation-break-line");
                    var dataIndex = 0;
                    for (var breakIdx = 0; breakIdx < breakCount; breakIdx++) {
                        // Header indices 0-9 (repeated for each break):
                        // [0]: Break Line Style (swBreakLineStyle_e: 1=Straight, 2=ZigZag, 3=Curve, 4=SmallZigZag, 5=Jagged)
                        // [1]: COLORREF (0 or -1 for default color)
                        // [2]: Line type (swLineTypes_e)
                        // [3]: Line style (swLineStyles_e)
                        // [4]: Line weight (swLineWeights_e)
                        // [5]: Layer ID
                        // [6]: Layer override (swLayerOverride_e)
                        // [7]: Number of line segments (for straight/zigzag)
                        // [8]: Number of arcs (for curve)
                        // [9]: Number of splines (for jagged)
                        var style = (swBreakLineStyle_e)breakData[dataIndex + 0];
                        var colorRef = (int)breakData[dataIndex + 1];
                        var lineType = (swLineTypes_e)breakData[dataIndex + 2];
                        var lineStyle = (swLineStyles_e)breakData[dataIndex + 3];
                        var lineWeight = (swLineWeights_e)breakData[dataIndex + 4];
                        var layerId = (int)breakData[dataIndex + 5];
                        var layerOverride = breakData[dataIndex + 6];
                        var numLines = (int)breakData[dataIndex + 7];
                        var numArcs = (int)breakData[dataIndex + 8];
                        var numSplines = (int)breakData[dataIndex + 9];
                        dataIndex += 10; // advance past header
                        var breakColor = (colorRef == 0 || colorRef == -1) ? "#000000" : ColorRefToHex(colorRef);
                        var breakStrokeWidth = MapLineWeight((int)lineWeight);
                        switch (style) {
                            case swBreakLineStyle_e.swBreakLine_Straight: // 2 lines * 1 segment * 2 points * 3 coords = 12 doubles
                            case swBreakLineStyle_e.swBreakLine_ZigZag: // 2 lines * 5 segments * 2 points * 3 coords = 60 doubles
                            case swBreakLineStyle_e.swBreakLine_SmallZigZag: // 2 lines * 5 segments * 2 points * 3 coords = 60 doubles
                                var pointsPerLine = 2;
                                for (var lineIdx = 0; lineIdx < numLines; lineIdx++) {
                                    var pathBuilder = new System.Text.StringBuilder();
                                    for (var ptIdx = 0; ptIdx < pointsPerLine; ptIdx++) {
                                        var x = breakData[dataIndex++];
                                        var y = breakData[dataIndex++];
                                        dataIndex++; // skip z
                                        var (svgX, svgY) = BreakLinePointToSvg(x, y);
                                        pathBuilder.Append(ptIdx == 0 ? $"M {Format(svgX)} {Format(svgY)}" : $" L {Format(svgX)} {Format(svgY)}");
                                        foreach (var writer in writers) writer.UpdateBounds(svgX, svgY);
                                    }
                                    foreach (var writer in writers) writer.AddPath(pathBuilder.ToString(), breakColor, breakStrokeWidth);
                                }
                                break;
                            case swBreakLineStyle_e.swBreakLine_Curve: // per arc: direction(1) + start(3) + end(3) + center(3) = 10 doubles
                                for (var arcIdx = 0; arcIdx < numArcs; arcIdx++) {
                                    var arcDirection = breakData[dataIndex++]; // 1 = CCW, -1 = CW
                                    var startX = breakData[dataIndex++];
                                    var startY = breakData[dataIndex++];
                                    dataIndex++; // skip startZ
                                    var endX = breakData[dataIndex++];
                                    var endY = breakData[dataIndex++];
                                    dataIndex++; // skip endZ
                                    var centerX = breakData[dataIndex++];
                                    var centerY = breakData[dataIndex++];
                                    dataIndex++; // skip centerZ
                                    var (svgStartX, svgStartY) = BreakLinePointToSvg(startX, startY);
                                    var (svgEndX, svgEndY) = BreakLinePointToSvg(endX, endY);
                                    var (svgCenterX, svgCenterY) = BreakLinePointToSvg(centerX, centerY);
                                    // Calculate radius from center to start point
                                    var radius = Math.Sqrt(Math.Pow(svgStartX - svgCenterX, 2) + Math.Pow(svgStartY - svgCenterY, 2));
                                    // SVG arc: A rx ry x-axis-rotation large-arc-flag sweep-flag x y
                                    // sweep-flag: 0 = CCW, 1 = CW (note: Y is flipped, so direction inverts)
                                    var sweepFlag = arcDirection > 0 ? 0 : 1;
                                    // Determine large-arc-flag based on arc angle
                                    var startAngle = Math.Atan2(svgStartY - svgCenterY, svgStartX - svgCenterX);
                                    var endAngle = Math.Atan2(svgEndY - svgCenterY, svgEndX - svgCenterX);
                                    var angleDiff = endAngle - startAngle;
                                    if (sweepFlag == 1 && angleDiff > 0) angleDiff -= 2 * Math.PI;
                                    if (sweepFlag == 0 && angleDiff < 0) angleDiff += 2 * Math.PI;
                                    var largeArcFlag = Math.Abs(angleDiff) > Math.PI ? 1 : 0;
                                    var arcPath = $"M {Format(svgStartX)} {Format(svgStartY)} A {Format(radius)} {Format(radius)} 0 {largeArcFlag} {sweepFlag} {Format(svgEndX)} {Format(svgEndY)}";
                                    foreach (var writer in writers) writer.AddPath(arcPath, breakColor, breakStrokeWidth);
                                    foreach (var writer in writers) writer.UpdateBounds(svgStartX, svgStartY);
                                    foreach (var writer in writers) writer.UpdateBounds(svgEndX, svgEndY);
                                }
                                break;
                            case swBreakLineStyle_e.swBreakLine_Jagged: // per spline: n(1 int) + 3n doubles
                                for (var splineIdx = 0; splineIdx < numSplines; splineIdx++) {
                                    var numPoints = (int)breakData[dataIndex++];
                                    if (numPoints < 2) continue;
                                    var pathBuilder = new System.Text.StringBuilder();
                                    for (var ptIdx = 0; ptIdx < numPoints; ptIdx++) {
                                        var x = breakData[dataIndex++];
                                        var y = breakData[dataIndex++];
                                        dataIndex++; // skip z
                                        var (svgX, svgY) = BreakLinePointToSvg(x, y);
                                        pathBuilder.Append(ptIdx == 0 ? $"M {Format(svgX)} {Format(svgY)}" : $" L {Format(svgX)} {Format(svgY)}");
                                        foreach (var writer in writers) writer.UpdateBounds(svgX, svgY);
                                    }
                                    foreach (var writer in writers) writer.AddPath(pathBuilder.ToString(), breakColor, breakStrokeWidth);
                                }
                                break;
                        }
                    }
                    foreach (var writer in writers) writer.EndGroup(); // End annotation-break-line
                }
            }
        }
        // ---------------------------------------------------------
        // Annotations
        // ---------------------------------------------------------
        (double svgX, double svgY) AnnotationToSvg(double x, double y) {
            var mmX = MeterToMM(x);
            var mmY = MeterToMM(y);
            (mmX, mmY) = applyTransformY(mmX, mmY);
            return (mmX, mmY);
        }
        var annotations = view.GetAnnotationsEx();
        foreach (var ann in annotations) {
            var annotationVisible = (swAnnotationVisibilityState_e)ann.Visible;
            if (annotationVisible != swAnnotationVisibilityState_e.swAnnotationVisible) continue;
            var isDangling = ann.IsDangling();
            var type = (swAnnotationType_e)ann.GetType();
            switch (type) {
                case swAnnotationType_e.swNote:
                    var pos = (double[])ann.GetPosition(); // x,y,z
                    var annX = pos[0];
                    var annY = pos[1];
                    // --
                    // GDUBE: it doesn't work.
                    // The break coords are still the same as when the view is un-break'd,
                    // while the annotation coords change when the view is break'd/un-break'd.
                    // Skip annotations that fall within the break region.
                    // if (hasXAxisBreak && annX.IsBetween(breakXAxis_x1, breakXAxis_x2)) continue;
                    // if (hasYAxisBreak && annY.IsBetween(breakYAxis_y1, breakYAxis_y2)) continue;
                    // --
                    if (pos == null) continue;
                    var (x, y) = AnnotationToSvg(annX, annY);
                    var text = "";
                    var colorRef = ann.Color;
                    var hexColor = ColorRefToHex(colorRef);
                    var note = (INote)ann.GetSpecificAnnotation();
                    if (!note.IsBomBalloon()) continue;
                    if (note != null) {
                        text = note.GetText();
                        var noteVisible = note.Visible;
                        var noteBehindSheet = note.BehindSheet;
                        var noteIsAttached = note.IsAttached();
                        // GDUBE: I used to skip annotations with "?", however I changed my mind for two reasons:
                        // 1. It displays visually that there's an issue with the slddrw & so a user could tell us to fix it.
                        // 2. It breaks grouped balloons if the first balloon is a "?" (in which case, skipping it would make a "hole" in the group & the arrow/line/lead would be missing).
                        // Skip annotations with "?" text and log a warning
                        // if (text == "?") {
                        //     warnings.Add($"Skipped balloon with '?' text at position ({x:F2}, {y:F2})");
                        //     continue;
                        // }
                        // if (isDangling) continue;

                        var balloonId = (!string.IsNullOrEmpty(text) && text.All(char.IsDigit))
                            ? $"balloon-{text}"
                            : $"balloon-{Guid.NewGuid().ToString().Split('-')[0]}";
                        var textCoordinates = NxPoint.FromCoords(x, y);
                        
                        // Build data attributes for the balloon group
                        var dataAttributes = new Dictionary<string, string> {
                            { "balloon-number", text },
                            { "content", text }
                        };
                        
                        // Add BOM metadata if available
                        if (includeBomMetadata && bomData.TryGetValue(text, out var bomItem)) {
                            if (!string.IsNullOrEmpty(bomItem.PartNumber)) {
                                dataAttributes["part-number"] = bomItem.PartNumber;
                            }
                            if (!string.IsNullOrEmpty(bomItem.Name)) {
                                dataAttributes["name"] = bomItem.Name;
                            }
                            if (!string.IsNullOrEmpty(bomItem.Specification)) {
                                dataAttributes["specification"] = bomItem.Specification;
                            }
                        }

                        var dData = (DisplayData)ann.GetDisplayData();
                        var dDataPlCount = dData.GetPolyLineCount();
                        var dDataEllipseCount = dData.GetEllipseCount();
                        var dDataLineCount = dData.GetLineCount();
                        var dDataPointCount = dData.GetPointCount();
                        var dDataParabolaCount = dData.GetParabolaCount();
                        var dDataArrowHeadCount = dData.GetArrowHeadCount();
                        var dDataArcCount = dData.GetArcCount();
                        var dDataPolygonCount = dData.GetPolygonCount();
                        var dDataTextCount = dData.GetTextCount();
                        var dDataTriangleCount = dData.GetTriangleCount();
                        // Skip if annotation text isn't displayed.
                        // Note: this works for hidden balloons when the view has (enabled/active) breaks (horizontal/vertical breaks).
                        // Note: works in Detailing Mode.
                        if (dDataTextCount == 0) continue;

                        // Start a group for this balloon annotation
                        foreach (var writer in writers) writer.StartGroup(balloonId, "annotation annotation-balloon", dataAttributes);
                        if (note.HasBalloon()) {
                            // --------------------------
                            // 1. Draw the balloon shape (circle)
                            // --------------------------
                            // [ centerX, centerY, centerZ, arcX, arcY, arcZ, radius ]
                            var balloonInfo = (double[])note.GetBalloonInfo();
                            var balloonCenter = NxPoint.FromSpan(balloonInfo.AsSpan().Slice(0, 3));
                            var balloonArc = NxPoint.FromSpan(balloonInfo.AsSpan().Slice(3, 3));
                            var balloonRadius = MeterToMM(balloonInfo[6]);
                            var style = (swBalloonStyle_e)note.GetBalloonStyle();
                            var (cx, cy) = AnnotationToSvg(balloonCenter.X, balloonCenter.Y);
                            if (style == swBalloonStyle_e.swBS_Circular) {
                                // SOLUTION (with a caveat):
                                // The balloon info returns coordinates in the sheet space.
                                // However, if the balloon is attached to a drawing view, the coordinates returned for the balloon will always be
                                // the coordinates of the balloon when the drawing was loaded (ie. file opened).
                                // If you move the view, and its annotations along with it, the coordinates returned will be wrong
                                // (ie. they will still be the coords of the balloon at the time the drawing was loaded).
                                foreach (var writer in writers)  writer.AddCircle(cx, cy, balloonRadius, hexColor, 0.25);
                                textCoordinates.X = cx; 
                                textCoordinates.Y = cy;
                                foreach(var writer in writers) writer.UpdateBounds(cx - balloonRadius, cy - balloonRadius);
                                foreach(var writer in writers) writer.UpdateBounds(cx + balloonRadius, cy + balloonRadius);
                                // WORKAROUND: (doesn't work very well however)
                                // foreach(var writer in writers) writer.AddCircle(x, y, balloonRadius - 0.8, hexColor, 0.25);
                            }
                            // TODO: Handle other styles (triangle, hexagon, etc.)
                            // --------------------------
                            // 2. Draw the leader line(s)
                            // --------------------------
                            var leaderCount = ann.GetLeaderCount();
                            for (var li = 0; li < leaderCount; li++) {
                                var leaderPoints = (double[])ann.GetLeaderPointsAtIndex(li);
                                if (leaderPoints != null && leaderPoints.Length >= 6) {
                                    // Draw line segments between consecutive points (x,y,z triplets)
                                    for (var pi = 0; pi < leaderPoints.Length - 3; pi += 3) {
                                        var (lx1, ly1) = AnnotationToSvg(leaderPoints[pi], leaderPoints[pi + 1]);
                                        var (lx2, ly2) = AnnotationToSvg(leaderPoints[pi + 3], leaderPoints[pi + 4]);
                                        foreach(var writer in writers) writer.AddLine(lx1, ly1, lx2, ly2, hexColor, 0.25);
                                        foreach(var writer in writers) writer.UpdateBounds(lx1, ly1);
                                        foreach(var writer in writers) writer.UpdateBounds(lx2, ly2);
                                    }
                                }
                            }
                            var multijogLeaderCount = ann.GetMultiJogLeaderCount();
                            if (multijogLeaderCount > 0) {
                                var multiJogLeadersArr = (object[])ann.GetMultiJogLeaders();
                                foreach (MultiJogLeader mjLeader in multiJogLeadersArr) {
                                    var lineCount = mjLeader.GetLineCount();
                                    for (var li = 0; li < lineCount; li++) {
                                        var lineData = (double[])mjLeader.GetLineAtIndex(li);
                                        var lineType = (swLineTypes_e)lineData[0];
                                        var startPoint = NxPoint.FromSpan(lineData.AsSpan().Slice(1, 3));
                                        var endPoint = NxPoint.FromSpan(lineData.AsSpan().Slice(4, 3));
                                        var (sx, sy) = AnnotationToSvg(startPoint.X, startPoint.Y);
                                        var (ex, ey) = AnnotationToSvg(endPoint.X, endPoint.Y);
                                        foreach (var writer in writers) writer.AddLine(sx, sy, ex, ey, hexColor, 0.25);
                                    }
                                }
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(text)) {
                            // ATTENTION: extent includes the leader (arrow), so it will be totally wrong.
                            // GetExtent: [lowerLeftX, lowerLeftY, lowerLeftZ, upperRightX, upperRightY, upperRightZ]
                            // in sheet space (meters). Use this to calculate the center of the text box.
                            // var extent = (double[])note.GetExtent();
                            // if (extent != null && extent.Length >= 6) {
                            //     var noteBottomLeft = NxPoint.FromSpan(extent.AsSpan().Slice(0, 3));
                            //     var noteUpperRight = NxPoint.FromSpan(extent.AsSpan().Slice(3, 3));
                            //     (x, y) = AnnotationToSvg(
                            //         ((noteUpperRight.X + noteBottomLeft.X) / 2),
                            //         ((noteUpperRight.Y + noteBottomLeft.Y) / 2)
                            //     );
                            // }
                            var fontSize = 8.0; // Default mm
                            var noteHeight = note.GetHeight();
                            var noteHeightPoints = note.GetHeightInPoints();
                            var fontFamily = "Arial";
                            var textFormat = (ITextFormat)ann.GetTextFormat(0);
                            if (textFormat != null) {
                                var typeFaceName = textFormat.TypeFaceName;
                                var lineLength = textFormat.LineLength;
                                var lineSpacing = textFormat.LineSpacing;
                                var charHeightInPts = textFormat.CharHeightInPts;
                                var charSpacingFactor = textFormat.CharSpacingFactor;
                                fontSize = textFormat.CharHeight * 1000.0; // Meters to MM
                            }
                            var lineHeight = fontSize * 1.2;
                            foreach (var writer in writers)  writer.AddText(textCoordinates.X, textCoordinates.Y, text, fontFamily, fontSize, hexColor, 0, "middle");
                            // Approximate text bounds (rough estimate based on font size and character count)
                            var textWidth = text.Length * fontSize * 0.6; // Rough approximation
                            var textHeight = fontSize * 1.2;
                            foreach(var writer in writers) writer.UpdateBounds(textCoordinates.X - textWidth / 2, textCoordinates.Y - textHeight / 2);
                            foreach(var writer in writers) writer.UpdateBounds(textCoordinates.X + textWidth / 2, textCoordinates.Y + textHeight / 2);
                        }
                        
                        // End the balloon annotation group
                        foreach(var writer in writers) writer.EndGroup();
                    }
                    break;
                default:
                    break;
            }
        }
    }

    /// <summary>
    /// Parses the flat double array returned by IView.GetPolylines6 into structured PolylineRecord objects.
    /// </summary>
    private List<NxPolylineRecord> ParsePolylines(double[] data) {
        var records = new List<NxPolylineRecord>();
        if (data == null || data.Length == 0) return records;
        var i = 0;
        while (i < data.Length) {
            // Check for trailing zeros or insufficient data for a valid record
            // Minimum record size: Type(1) + GeomDataSize(1) + LineColor(1) + LineStyle(1) + LineFont(1) + LineWeight(1) + LayerID(1) + LayerOverride(1) + NumPolyPoints(1) = 9
            if (i + 9 > data.Length) break;
            // Check if we've hit trailing zeros (Type would be 0 or 1, not some random value)
            var typeValue = data[i];
            if (typeValue != 0 && typeValue != 1) break;
            var record = new NxPolylineRecord();
            record.Type = (NxPolylineType)(int)data[i++];
            var geomDataSize = (int)data[i++];
            // Handle the GeomDataSize = 17 bug: treat as 0 (trailing zeros)
            if (geomDataSize == 17) geomDataSize = 0;
            // Validate we have enough data for GeomData + metadata
            if (i + geomDataSize + 7 > data.Length) break;
            // Read GeomData if present (for arcs: 12 values)
            if (geomDataSize > 0) {
                record.GeomData = new double[geomDataSize];
                Array.Copy(data, i, record.GeomData, 0, geomDataSize);
                i += geomDataSize;
            }
            // Read metadata fields
            record.LineColor = (int)data[i++];
            record.LineStyle = (int)data[i++];
            record.LineFont = (int)data[i++];
            record.LineWeight = data[i++];
            record.LayerID = (int)data[i++];
            record.LayerOverride = (int)data[i++];
            // Validate we have NumPolyPoints
            if (i >= data.Length) break;
            var numPolyPoints = (int)data[i++];
            // Validate we have enough data for all points
            var pointsDataSize = numPolyPoints * 3;
            if (i + pointsDataSize > data.Length) break;
            // Read tessellated points
            record.Points = new List<NxPoint>(numPolyPoints);
            for (var p = 0; p < numPolyPoints; p++) {
                var x = data[i++];
                var y = data[i++];
                var z = data[i++];
                record.Points.Add(new NxPoint(x, y, z));
            }
            records.Add(record);
        }
        return records;
    }

    /// <summary>
    /// Convert line weight value to stroke width in mm.
    /// The values in inches were taken from Solidworks' defaults.
    /// </summary>
    private double MapLineWeight(int value) {
        const double inchToMm = 25.4;
        switch ((swLineWeights_e)value) {
            case swLineWeights_e.swLW_THIN:    return 0.0071 * inchToMm;
            case swLineWeights_e.swLW_NORMAL:  return 0.0098 * inchToMm;
            case swLineWeights_e.swLW_THICK:   return 0.0138 * inchToMm;
            case swLineWeights_e.swLW_THICK2:  return 0.0197 * inchToMm;
            case swLineWeights_e.swLW_THICK3:  return 0.0276 * inchToMm;
            case swLineWeights_e.swLW_THICK4:  return 0.0394 * inchToMm;
            case swLineWeights_e.swLW_THICK5:  return 0.0551 * inchToMm;
            case swLineWeights_e.swLW_THICK6:  return 0.0787 * inchToMm;
            default: return 0.0098 * inchToMm; // Default to Normal
        }
    }

    /// <summary>
    /// Converts COLORREF (0x00BBGGRR) to string hex representation (eg. "#000000").
    /// </summary>
    private string ColorRefToHex(int colorRef) {
        if (colorRef == -1) return "#000000";
        var r = colorRef & 0xFF;
        var g = (colorRef >> 8) & 0xFF;
        var b = (colorRef >> 16) & 0xFF;
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private string Format(double value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>
    /// Parses the BOM table from a drawing and returns a dictionary mapping item numbers to BOM data.
    /// Supports French BOM headers: ITEM, QTE, # PIÈCE, NOM, SPÉCIFICATION
    /// </summary>
    private Dictionary<string, BomItem> ParseBomTable(DrawingDoc drawing) {
        var bomData = new Dictionary<string, BomItem>();
        
        try {
            // Get the first sheet
            var sheet = (Sheet)drawing.GetCurrentSheet();
            if (sheet == null) return bomData;
            
            // Get table annotations from the sheet
            var view = (IView)drawing.GetFirstView(); // Sheet view
            if (view == null) return bomData;
            while (view != null) {
                var annotations = (object[])view.GetTableAnnotations();
                if (annotations != null) {
                    foreach (var ann in annotations) {
                        var tableAnn = ann as ITableAnnotation;
                        if (tableAnn == null) continue;
                        
                        // Check if this is a BOM table
                        if (tableAnn.Type != (int)swTableAnnotationType_e.swTableAnnotation_BillOfMaterials) {
                            continue;
                        }
                        
                        // Parse the BOM table
                        var rowCount = tableAnn.RowCount;
                        var colCount = tableAnn.ColumnCount;
                        
                        if (rowCount < 2 || colCount < 1) continue;
                        
                        // Find column indices by examining the header row
                        var itemColIndex = -1;
                        var partNumberColIndex = -1;
                        var nameColIndex = -1;
                        var specificationColIndex = -1;
                        
                        for (var col = 0; col < colCount; col++) {
                            var headerText = tableAnn.DisplayedText2[0, col, true];
                            if (string.IsNullOrEmpty(headerText)) continue;
                            
                            var header = headerText.Trim().ToUpperInvariant();
                            
                            if (header.Contains("ITEM") || header == "REPÈRE") {
                                itemColIndex = col;
                            } else if (header.Contains("PIÈCE") || header.Contains("PIECE") || header == "# PIÈCE") {
                                partNumberColIndex = col;
                            } else if (header.Contains("NOM") || header.Contains("DESCRIPTION")) {
                                nameColIndex = col;
                            } else if (header.Contains("SPÉCIFICATION") || header.Contains("SPECIFICATION") || header.Contains("SPEC")) {
                                specificationColIndex = col;
                            }
                        }
                        
                        // If we didn't find the item column, we can't map the data
                        if (itemColIndex == -1) continue;
                        
                        // Parse data rows (skip header row 0)
                        for (var row = 1; row < rowCount; row++) {
                            var itemNumber = tableAnn.DisplayedText2[row, itemColIndex, true]?.Trim();
                            if (string.IsNullOrEmpty(itemNumber)) continue;
                            
                            var bomItem = new BomItem {
                                ItemNumber = itemNumber,
                                PartNumber = partNumberColIndex >= 0 ? tableAnn.DisplayedText2[row, partNumberColIndex, true]?.Trim() : null,
                                Name = nameColIndex >= 0 ? tableAnn.DisplayedText2[row, nameColIndex, true]?.Trim() : null,
                                Specification = specificationColIndex >= 0 ? tableAnn.DisplayedText2[row, specificationColIndex, true]?.Trim() : null
                            };
                            
                            // Store in dictionary (use item number as key)
                            if (!bomData.ContainsKey(itemNumber)) {
                                bomData[itemNumber] = bomItem;
                            }
                        }
                        
                        // Found and parsed a BOM table, return the data
                        return bomData;
                    }
                }
                
                view = (IView)view.GetNextView();
            }
        } catch {
            // Silent failure - return empty dictionary
        }
        
        return bomData;
    }

    /// <summary>
    /// Maps SolidWorks text justification to SVG text-anchor attribute.
    /// </summary>
    private static string MapJustificationToTextAnchor(swTextJustification_e justification) {
        return justification switch {
            swTextJustification_e.swTextJustificationLeft => "start",
            swTextJustification_e.swTextJustificationCenter => "middle",
            swTextJustification_e.swTextJustificationRight => "end",
            _ => "start" // swTextJustificationNone defaults to start
        };
    }

    /// <summary>
    /// Maps SolidWorks text reference position to SVG dominant-baseline attribute.
    /// The reference position indicates which point of the text bounding box the coordinates refer to.
    /// </summary>
    private static (string dominantBaseline, string textAnchor) MapRefPositionToBaseline(swTextPosition_e refPosition) {
        // TextAnchor: start, middle, end
        // DominantBaseline:
        // - auto (bottom)
        // - middle (middle on y-axis)
        // - hanging (top of text)
        // - text-top (like auto)
        return refPosition switch {
            swTextPosition_e.swUPPER_LEFT => ("hanging", "start"),
            swTextPosition_e.swUPPER_RIGHT => ("hanging", "end"),
            swTextPosition_e.swUPPER_CENTER => ("hanging", "middle"),
            swTextPosition_e.swCENTER => ("middle", "middle"),
            swTextPosition_e.swLOWER_LEFT => ("auto", "start"),
            swTextPosition_e.swLOWER_RIGHT => ("auto", "end"),
            _ => ("auto", "start")
        };
    }

    private static void DebugAnnotation(Annotation ann) {
        var annType = (swAnnotationType_e)ann.GetType();
        if (annType != swAnnotationType_e.swNote) return;
        var note = (INote)ann.GetSpecificAnnotation();
        if (note is null) return;
        var isBomBalloon = note.IsBomBalloon();
        var visible = (swAnnotationVisibilityState_e)ann.Visible;
        var noteText = note.GetText();
        var noteName = note.GetName();
        var tagName = note.TagName;
        var promptText = note.PromptText;
        var propertyLinkedText = note.PropertyLinkedText;
        var textCount = note.GetTextCount();
        if (textCount > 1 && noteText.Trim([' ']) == "\r\n") {
            var lines = new string[textCount];
            for (var i = 1; i <= textCount; i++) {
                var textAtI = note.GetTextAtIndex(i);
                var textHeight = note.GetTextHeightAtIndex(i) * 1000.0; // Height in MM (aka. font size)
                var textLinespacing = note.GetTextLineSpacingAtIndex(i) * 1000.0; // Line spacing in MM
                var textJustificationAtIndex = (swTextJustification_e)note.GetTextJustificationAtIndex(i);
                var textPositionAtIndex = note.GetTextPositionAtIndex(i); // Position relative to the sheet origin.
                var textRefPointAtIndex = (swTextPosition_e)note.GetTextRefPositionAtIndex(i); // Reference position: upper left, lower left, center
                lines[i - 1] = textAtI;
            }
            noteText = string.Join(noteText, lines);
        }
        var position = (double[])ann.GetPosition();
        var textJustification = (swTextJustification_e)note.GetTextJustification();
        var textVerticalJustification = (swTextAlignmentVertical_e)note.GetTextVerticalJustification();
        var height = note.GetHeight() * 1000.0;
        var textPoint = note.IGetTextPoint2(); // Note origin. Relative to ?
        var textUpperRight = (double[])note.GetUpperRight(); // Top-right point relative to sheet origin
        var textFormat = (ITextFormat)ann.GetTextFormat(0);
        var typeFaceName = textFormat.TypeFaceName; // font family name
        var lineLength = textFormat.LineLength; // text line length in meters (always zero?)
        var lineSpacing = textFormat.LineSpacing; // text line spacing
        var lineSpacingMM = lineSpacing * 1000.0; // Meters to MM
        var charSpacingFactor = textFormat.CharSpacingFactor; // spacing between characters
        var charHeightInPts = textFormat.CharHeightInPts; // height of the font in points
        var charHeight = textFormat.CharHeight; // height of the font in meters
        var charHeightMM = charHeight * 1000.0; // Meters to MM
        var italic = textFormat.Italic;
        var bold = textFormat.Bold;
        var escapement = textFormat.Escapement; // text angle in radians
        var obliqueAngle = textFormat.ObliqueAngle; // text slant
        var strikeout = textFormat.Strikeout;
        var underline = textFormat.Underline;
        var widthFactor = textFormat.WidthFactor; // Stretches the text by the specified factor.  
        var backwards = textFormat.BackWards;
        var vertical = textFormat.Vertical; // individual characters as vertical
    }
}
