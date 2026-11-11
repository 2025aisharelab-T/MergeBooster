using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

namespace IcoGenerator
{
    class Program
    {
        static void Main()
        {
            using (Bitmap bmp = new Bitmap(64, 64))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);

                    // 1. Engine Flame (Vibrant Orange & Yellow)
                    using (SolidBrush flame1 = new SolidBrush(ColorTranslator.FromHtml("#FF4500")))
                        g.FillPolygon(flame1, new Point[] { new Point(14, 48), new Point(4, 60), new Point(22, 54) });
                    using (SolidBrush flame2 = new SolidBrush(ColorTranslator.FromHtml("#FFD700")))
                        g.FillPolygon(flame2, new Point[] { new Point(16, 48), new Point(8, 56), new Point(20, 52) });

                    // 2. Fins (Vibrant Coral Red)
                    using (SolidBrush wing = new SolidBrush(ColorTranslator.FromHtml("#EF4444")))
                    {
                        g.FillPolygon(wing, new Point[] { new Point(14, 28), new Point(2, 44), new Point(22, 40) });
                        g.FillPolygon(wing, new Point[] { new Point(36, 14), new Point(50, 2), new Point(46, 22) });
                    }

                    // 3. Rocket Capsule Body (Pure White & Slate Shadow)
                    using (GraphicsPath bodyPath = new GraphicsPath())
                    {
                        bodyPath.AddArc(12, 8, 40, 40, 225, 180);
                        bodyPath.AddLine(48, 36, 24, 48);
                        bodyPath.CloseFigure();

                        using (SolidBrush bBrush = new SolidBrush(ColorTranslator.FromHtml("#FFFFFF")))
                            g.FillPath(bBrush, bodyPath);
                        using (Pen pen = new Pen(ColorTranslator.FromHtml("#CBD5E1"), 2f))
                            g.DrawPath(pen, bodyPath);
                    }

                    // 4. Red Rocket Nose Cone Tip
                    using (SolidBrush nose = new SolidBrush(ColorTranslator.FromHtml("#DC2626")))
                    {
                        g.FillPie(nose, 28, 4, 28, 28, 270, 90);
                    }

                    // 5. Porthole Glass Window (Cyan Porthole & White Glass Specular Highlight)
                    using (SolidBrush ring = new SolidBrush(ColorTranslator.FromHtml("#1E293B")))
                        g.FillEllipse(ring, 24, 16, 20, 20);
                    using (SolidBrush glass = new SolidBrush(ColorTranslator.FromHtml("#38BDF8")))
                        g.FillEllipse(glass, 27, 19, 14, 14);
                    using (SolidBrush spec = new SolidBrush(Color.White))
                        g.FillEllipse(spec, 30, 21, 5, 5);
                }

                IntPtr hIcon = bmp.GetHicon();
                using (Icon icon = Icon.FromHandle(hIcon))
                {
                    using (FileStream fs = new FileStream("app_icon.ico", FileMode.Create))
                    {
                        icon.Save(fs);
                    }
                }
            }
            Console.WriteLine("app_icon.ico generated successfully!");
        }
    }
}
