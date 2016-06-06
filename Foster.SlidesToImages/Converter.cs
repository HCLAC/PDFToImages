using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;

namespace Foster.SlidesToImages
{
    public class Converter
    {
        private String PathToDll { get; set; }
        public Size NewSize { get; set; }

        public Converter(string pathToLibrary = "")
        {
            if (!String.IsNullOrEmpty(pathToLibrary))
            {
                PathToDll = pathToLibrary;
            }
            else
            {
                PathToDll = System.Web.HttpContext.Current.Server.MapPath("~/FosterSlidesToImagesLib/");
            }
        }

        /// <summary>
        /// Convert specified PDF file to Images (1 image per pdf page)
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <returns></returns>
        public IEnumerable<Image> ConvertPDFToImages(string pdfPath)
        {
            IEnumerable<Image> retorno = new List<Image>();
            var task = Task.Factory.StartNew(() =>
            {
                retorno = ExtractImages(pdfPath);
            });
            task.Wait();
            return retorno;
        }

        /// <summary>
        /// Convert specified PDF file to Images (1 image per pdf page)
        /// </summary>
        /// <param name="pdfStream"></param>
        /// <returns></returns>
        public IEnumerable<Image> ConvertPDFToImages(Stream pdfStream)
        {
            IEnumerable<Image> retorno = new List<Image>();
            var task = Task.Factory.StartNew(() =>
            {
                retorno = ExtractImages(pdfStream);
            });
            task.Wait();
            return retorno;
        }

        /// <summary>
        /// Convert specified PDF file to Images (1 image per pdf page)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        IEnumerable<Image> ExtractImages(string file)
        {
            Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer = null;
            Ghostscript.NET.GhostscriptVersionInfo vesion;
            vesion = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0),
                        PathToDll + @"\gsdll32.dll",
                        string.Empty,
                        Ghostscript.NET.GhostscriptLicense.GPL);

            using (rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
            {
                try
                {
                    rasterizer.Open(file, vesion, false);
                    return GetImagesFromRasterizer(rasterizer);
                }
                catch
                {
                    vesion = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0),
                        PathToDll + @"\gsdll64.dll",
                        string.Empty,
                        Ghostscript.NET.GhostscriptLicense.GPL);

                    rasterizer.Open(file, vesion, false);
                    return GetImagesFromRasterizer(rasterizer);
                }
                
            }
        }

        /// <summary>
        /// Convert specified PDF file to Images (1 image per pdf page)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        IEnumerable<Image> ExtractImages(Stream pdfStream)
        {
            Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer = null;
            Ghostscript.NET.GhostscriptVersionInfo vesion = null;
            vesion = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0),
                        PathToDll + @"\gsdll32.dll",
                        string.Empty,
                        Ghostscript.NET.GhostscriptLicense.GPL);
            using (rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
            {
                try
                {
                    rasterizer.Open(pdfStream, vesion, false);
                    return GetImagesFromRasterizer(rasterizer);
                }
                catch
                {
                    vesion = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0),
                        PathToDll + @"\gsdll64.dll",
                        string.Empty,
                        Ghostscript.NET.GhostscriptLicense.GPL);

                    rasterizer.Open(pdfStream, vesion, false);
                    return GetImagesFromRasterizer(rasterizer);
                }
            }
        }

        /// <summary>
        /// Get the images from PDF files in rasterizer
        /// </summary>
        /// <param name="rasterizer"></param>
        /// <returns></returns>
        IEnumerable<Image> GetImagesFromRasterizer(Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer)
        {
            var imagesList = new List<Image>();
            var task = Task.Factory.StartNew(() => {
                for (int i = 1; i <= rasterizer.PageCount; i++)
                {
                    Image img = rasterizer.GetPage(300, 300, i);

                    EncoderParameter qualityParam =
                        new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 60L);

                    EncoderParameters encoderParams = new EncoderParameters
                    {
                        Param = new EncoderParameter[] { qualityParam }
                    };

                    var imageStream = new MemoryStream();
                    img.Save(imageStream, GetEncoderInfo("image/jpeg"), encoderParams);

                    Image imageExported = new Bitmap(imageStream);
                    if (NewSize != null)
                    {
                        imageExported = Util.ResizeImage(imageExported, NewSize);
                    }
                    imagesList.Add(imageExported);
                }
            });
            task.Wait();
            rasterizer.Close();
            return imagesList;
        }

        /// <summary> 
        /// Returns the image codec with the given mime type 
        /// </summary> 
        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            // Get image codecs for all image formats 
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

            // Find the correct image codec 
            for (int i = 0; i < codecs.Length; i++)
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];

            return null;
        }
    }
}
