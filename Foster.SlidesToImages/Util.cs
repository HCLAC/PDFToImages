using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Foster.SlidesToImages
{
    public class Util
    {
        /// <summary>
        /// Resizes the image to the specified size keeping the aspect ratio
        /// </summary>
        /// <param name="imgToResize"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public static Image ResizeImage(Image imgToResize, Size size)
        {
            //Get the image current width
            int sourceWidth = imgToResize.Width;
            //Get the image current height
            int sourceHeight = imgToResize.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            if (size.Width > 0)
            {
                //Calulate  width with new desired size
                nPercentW = ((float)size.Width / (float)sourceWidth);
            }
            if (size.Height > 0)
            {
                //Calculate height with new desired size
                nPercentH = ((float)size.Height / (float)sourceHeight);
            }

            if (nPercentH < nPercentW && nPercentH > 0)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;

            //New Width
            int destWidth = (int)(sourceWidth * nPercent);
            //New Height
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            // Draw image with new width and height
            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();

            return (Image)b;
        }
    }
}
