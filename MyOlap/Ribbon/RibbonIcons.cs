using Svg;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;

namespace MyOlap.Ribbon;

/// <summary>
/// Renders ribbon icons by loading the embedded Fluent SVG files via Svg.NET.
/// currentColor is resolved to Office icon grey (#212121) before rendering.
/// </summary>
internal static class RibbonIcons
{
    private const int Size = 32;
    private const string IconColor = "#212121";

    public static Bitmap MoveToRow()    => Render("MyOlap.Ribbon.move-to-row-fluent.svg");
    public static Bitmap MoveToCol()    => Render("MyOlap.Ribbon.move-to-column-fluent.svg");
    public static Bitmap MoveToHeader() => Render("MyOlap.Ribbon.move-to-header-fluent.svg");

    static Bitmap Render(string resourceName)
    {
        var asm = Assembly.GetExecutingAssembly();
        using var stream = asm.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Embedded SVG resource not found: {resourceName}");
        using var reader = new StreamReader(stream);
        var svg = reader.ReadToEnd().Replace("currentColor", IconColor);
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(svg));
        var doc = SvgDocument.Open<SvgDocument>(ms);
        return doc.Draw(Size, Size);
    }
}
