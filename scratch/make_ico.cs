using System;
using System.Drawing;
using System.IO;

namespace MakeIco
{
    class Program
    {
        static void Main()
        {
            string pngPath = "rocket_hd.png";
            if (!File.Exists(pngPath)) pngPath = "rocket.png";

            using (Bitmap orig = new Bitmap(pngPath))
            {
                // Make a clean 48x48 icon bitmap
                using (Bitmap bmp = new Bitmap(48, 48))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.Clear(Color.Transparent);
                        g.DrawImage(orig, 0, 0, 48, 48);
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
            }
            Console.WriteLine("app_icon.ico created from user rocket PNG successfully!");
        }
    }
}
