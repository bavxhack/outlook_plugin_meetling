using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Meetling.OutlookAddIn;

internal static class RibbonImages
{
    // Keep the bitmap alive for the OLE picture's lifetime. Load only when Office asks for an image.
    private static Bitmap? bitmap;
    private static object? picture;
    private static readonly object Gate = new object();

    internal static object GetPicture()
    {
        lock (Gate)
        {
            if (picture != null) return picture;
            var candidate = RenderCube(64);
            try { picture = PictureBridge.Convert(candidate); }
            catch { candidate.Dispose(); throw; }
            bitmap = candidate;
            return picture;
        }
    }

    internal static Bitmap RenderCube(int size)
    {
        using (var stream = typeof(RibbonImages).Assembly.GetManifestResourceStream(
            "Meetling.OutlookAddIn.Resources.MeetlingCube.svg"))
        {
            if (stream == null) throw new InvalidOperationException("Embedded Meetling cube missing.");
            var root = XDocument.Load(stream).Root!;
            XNamespace ns = "http://www.w3.org/2000/svg";
            var rect = root.Element(ns + "rect")!;
            var result = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                using (var graphics = Graphics.FromImage(result))
                {
                    graphics.Clear(Color.Transparent);
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.ScaleTransform(size / Number(rect, "width"), size / Number(rect, "height"));
                    var width = Number(rect, "width");
                    var height = Number(rect, "height");
                    var diameter = Number(rect, "rx") * 2;
                    using (var path = new GraphicsPath())
                    using (var fill = new SolidBrush(ColorTranslator.FromHtml((string)rect.Attribute("fill")!)))
                    {
                        path.AddArc(0, 0, diameter, diameter, 180, 90);
                        path.AddArc(width - diameter, 0, diameter, diameter, 270, 90);
                        path.AddArc(width - diameter, height - diameter, diameter, diameter, 0, 90);
                        path.AddArc(0, height - diameter, diameter, diameter, 90, 90);
                        path.CloseFigure();
                        graphics.FillPath(fill, path);
                    }
                    foreach (var circle in root.Elements(ns + "circle"))
                    {
                        var radius = Number(circle, "r");
                        using (var fill = new SolidBrush(ColorTranslator.FromHtml((string)circle.Attribute("fill")!)))
                            graphics.FillEllipse(fill, Number(circle, "cx") - radius,
                                Number(circle, "cy") - radius, radius * 2, radius * 2);
                    }
                }
                return result;
            }
            catch { result.Dispose(); throw; }
        }
    }

    private static float Number(XElement element, string name)
        => float.Parse((string)element.Attribute(name)!, CultureInfo.InvariantCulture);

    private sealed class PictureBridge : AxHost
    {
        private PictureBridge() : base(string.Empty) { }
        internal static object Convert(Image image) => GetIPictureDispFromPicture(image);
    }
}
