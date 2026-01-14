using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;

namespace Dx.Sw.SvgExport.Command;

public class SvgWriter {
    public string DocumentName { get; set; }

    private readonly XElement _svg;
    private readonly XElement _contentGroup;
    private readonly XNamespace _ns = "http://www.w3.org/2000/svg";
    private readonly Stack<XElement> _groupStack = new Stack<XElement>();

    // Bounds tracking for fit-to-content
    private double _minX = double.MaxValue;
    private double _minY = double.MaxValue;
    private double _maxX = double.MinValue;
    private double _maxY = double.MinValue;

    public double Width { get; }
    public double Height { get; }

    public SvgWriter(double width, double height) {
        Width = width;
        Height = height;
        _svg = new XElement(_ns + "svg",
            new XAttribute("width", $"{width}"),
            new XAttribute("height", $"{height}"),
            new XAttribute("viewBox", $"0 0 {width} {height}"),
            // For debugging:
            // new XAttribute("style", $"border: 1px solid black"),
            new XAttribute("xmlns", _ns.NamespaceName)
        );
        // Content group for all SVG elements
        _contentGroup = new XElement(_ns + "g");
        _svg.Add(_contentGroup);
    }

    private XElement GetCurrentContainer() {
        return _groupStack.Count > 0 ? _groupStack.Peek() : _contentGroup;
    }

    public void AddPath(string pathData, string strokeColor, double strokeWidth, string fillColor = "none", double[] strokeDashArray = null) {
        var path = new XElement(_ns + "path",
            new XAttribute("d", pathData),
            new XAttribute("stroke", strokeColor),
            new XAttribute("stroke-width", strokeWidth),
            new XAttribute("fill", fillColor)
        );
        if (strokeDashArray != null) {
            path.Add(
                new XAttribute("stroke-dasharray", string.Join(" ", strokeDashArray))
            );
        }
        GetCurrentContainer().Add(path);
    }
    
    public void AddLine(double x1, double y1, double x2, double y2, string strokeColor, double strokeWidth) {
        var line = new XElement(_ns + "line",
            new XAttribute("x1", Format(x1)),
            new XAttribute("y1", Format(y1)),
            new XAttribute("x2", Format(x2)),
            new XAttribute("y2", Format(y2)),
            new XAttribute("stroke", strokeColor),
            new XAttribute("stroke-width", strokeWidth)
        );
        GetCurrentContainer().Add(line);
    }

    public void AddCircle(double cx, double cy, double r, string strokeColor, double strokeWidth, string fillColor = "none") {
        var circle = new XElement(_ns + "circle",
            new XAttribute("cx", Format(cx)),
            new XAttribute("cy", Format(cy)),
            new XAttribute("r", Format(r)),
            new XAttribute("stroke", strokeColor),
            new XAttribute("stroke-width", strokeWidth),
            new XAttribute("fill", fillColor)
        );
        GetCurrentContainer().Add(circle);
    }

    public void AddEllipse(double cx, double cy, double rx, double ry, double rotation, string strokeColor, double strokeWidth, string fillColor = "none") {
        var ellipse = new XElement(_ns + "ellipse",
            new XAttribute("cx", Format(cx)),
            new XAttribute("cy", Format(cy)),
            new XAttribute("rx", Format(rx)),
            new XAttribute("ry", Format(ry)),
            new XAttribute("stroke", strokeColor),
            new XAttribute("stroke-width", strokeWidth),
            new XAttribute("fill", fillColor)
        );

        if (rotation != 0) {
            ellipse.Add(new XAttribute("transform", $"rotate({Format(rotation)}, {Format(cx)}, {Format(cy)})"));
        }

        GetCurrentContainer().Add(ellipse);
    }

    public void AddText(
        double x, 
        double y, 
        string text, 
        string fontFamily, 
        double fontSize, 
        string color, 
        double rotation = 0, 
        string textAnchor = "start", 
        string dominantBaseline = "middle",
        bool bold = false,
        bool italic = false
    ) {
        var textElement = new XElement(_ns + "text",
            new XAttribute("x", Format(x)),
            new XAttribute("y", Format(y)),
            new XAttribute("font-family", fontFamily),
            new XAttribute("font-size", fontSize),
            new XAttribute("fill", color),
            new XAttribute("text-anchor", textAnchor),
            new XAttribute("dominant-baseline", dominantBaseline)
        );
        if (bold) {
            textElement.Add(new XAttribute("font-weight", "bold"));
        }
        if (italic) {
            textElement.Add(new XAttribute("font-style", "italic"));
        }
        // Split by common newline patterns
        var lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        if (lines.Length == 1) {
            // Single line - just set text content
            textElement.Value = text;
        } else {
            // Multi-line - use tspan elements
            for (int i = 0; i < lines.Length; i++) {
                var tspan = new XElement(_ns + "tspan",
                    new XAttribute("x", Format(x)),
                    new XAttribute("dy", i == 0 ? "0" : "1.2em"),
                    lines[i]
                );
                textElement.Add(tspan);
            }
        }
        if (rotation != 0) {
            textElement.Add(new XAttribute("transform", $"rotate({Format(rotation)}, {Format(x)}, {Format(y)})"));
        }
        GetCurrentContainer().Add(textElement);
    }

    public void StartGroup(string id = null, string className = null, Dictionary<string, string> dataAttributes = null) {
        var group = new XElement(_ns + "g");
        
        if (!string.IsNullOrEmpty(id)) {
            group.Add(new XAttribute("id", id));
        }
        
        if (!string.IsNullOrEmpty(className)) {
            group.Add(new XAttribute("class", className));
        }
        
        if (dataAttributes != null) {
            foreach (var kvp in dataAttributes) {
                group.Add(new XAttribute("data-" + kvp.Key, kvp.Value));
            }
        }
        
        GetCurrentContainer().Add(group);
        _groupStack.Push(group);
    }

    public void EndGroup() {
        if (_groupStack.Count > 0) {
            _groupStack.Pop();
        }
    }

    public void AddTitle(string text) {
        if (string.IsNullOrEmpty(text)) return;
        
        var title = new XElement(_ns + "title", text);
        GetCurrentContainer().Add(title);
    }

    public void SetViewBox(double minX, double minY, double width, double height) {
        var viewBoxAttr = _svg.Attribute("viewBox");
        if (viewBoxAttr != null) {
            viewBoxAttr.Value = $"{Format(minX)} {Format(minY)} {Format(width)} {Format(height)}";
        }
    }

    public void SetDimensions(double width, double height) {
        void SetAttr(string attrName, double value) {
            var attr = _svg.Attribute(attrName);
            if (attr != null) {
                attr.Value = $"{Format(value)}";
            }
        }
        SetAttr("width", width);
        SetAttr("height", height);
    }

    /// <summary>
    /// Updates the bounds tracking with a new point.
    /// </summary>
    public void UpdateBounds(double x, double y) {
        if (x < _minX) _minX = x;
        if (x > _maxX) _maxX = x;
        if (y < _minY) _minY = y;
        if (y > _maxY) _maxY = y;
    }

    /// <summary>
    /// Applies fit-to-content transformation by adjusting the viewBox and dimensions
    /// to fit the tracked bounds with the specified padding.
    /// </summary>
    public void ApplyFitToContent(double padding = 5.0) {
        if (_minX == double.MaxValue || _maxX == double.MinValue) {
            // No bounds were tracked, nothing to fit
            return;
        }

        var contentMinX = _minX - padding;
        var contentMinY = _minY - padding;
        var contentWidth = (_maxX - _minX) + (2 * padding);
        var contentHeight = (_maxY - _minY) + (2 * padding);
        
        SetViewBox(contentMinX, contentMinY, contentWidth, contentHeight);
        SetDimensions(contentWidth, contentHeight);
    }

    public void Save(string filePath) {
        var doc = new XDocument(_svg);
        doc.Save(filePath);
    }

    private string Format(double value) {
        return value.ToString("0.####", CultureInfo.InvariantCulture);
    }
}
