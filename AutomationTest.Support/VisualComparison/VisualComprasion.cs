using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace AutomationUtils.VisualComparison
{
    public class VisualComprasion
    {
        public static void ScreenCapture(string filePath)
        {
            using (Bitmap bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bmpScreenshot))
                {
                    g.CopyFromScreen(Screen.PrimaryScreen.Bounds.X,
                                     Screen.PrimaryScreen.Bounds.Y,
                                     0,
                                     0,
                                     Screen.PrimaryScreen.Bounds.Size,
                                     CopyPixelOperation.SourceCopy);

                    bmpScreenshot.Save(filePath, ImageFormat.Png);
                }
            }
        }

        public static bool ImageCompareString(string filePath1, string filePath2 )
        {
            Bitmap bmp1 = new Bitmap(filePath1);
            Bitmap bmp2 = new Bitmap(filePath2);

            bool equals = true;
            bool flag = true;  //Inner loop isn't broken

            //Test to see if we have the same size of image
            if (bmp1.Size == bmp2.Size)
            {
                for (int x = 0; x < bmp1.Width; ++x)
                {
                    for (int y = 0; y < bmp1.Height; ++y)
                    {
                        if (bmp1.GetPixel(x, y) != bmp2.GetPixel(x, y))
                        {
                            equals = false;
                            flag = false;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        break;
                    }
                }
            }
            else
            {
                equals = false;
            }
            return equals;
        }
    }

}
