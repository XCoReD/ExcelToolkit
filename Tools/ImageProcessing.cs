using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Reflection;
using System.Drawing.Imaging;
using System.IO;

namespace Tools
{
    public static class ImageProcessing
    {
        //https://stackoverflow.com/questions/1397512/find-image-format-using-bitmap-object-in-c-sharp
        public static System.Drawing.Imaging.ImageFormat GetImageFormat(this System.Drawing.Image img)
        {
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Jpeg))
                return System.Drawing.Imaging.ImageFormat.Jpeg;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Bmp))
                return System.Drawing.Imaging.ImageFormat.Bmp;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Png))
                return System.Drawing.Imaging.ImageFormat.Png;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Emf))
                return System.Drawing.Imaging.ImageFormat.Emf;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Exif))
                return System.Drawing.Imaging.ImageFormat.Exif;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Gif))
                return System.Drawing.Imaging.ImageFormat.Gif;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Icon))
                return System.Drawing.Imaging.ImageFormat.Icon;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.MemoryBmp))
                return System.Drawing.Imaging.ImageFormat.MemoryBmp;
            if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Tiff))
                return System.Drawing.Imaging.ImageFormat.Tiff;
            else
                return System.Drawing.Imaging.ImageFormat.Wmf;
        }

        //may help in some cases, but in general makes things worse than painting by default
        public static Image FromWmf(Stream stream)
        {
            using (Metafile img = new Metafile(stream))
            {
                MetafileHeader metafileHeader = img.GetMetafileHeader();
                float scale = metafileHeader.DpiX / 96f;
                Bitmap bitmap = new Bitmap(
                    (int)(scale * img.Width / metafileHeader.DpiX * 100),
                    (int)(scale * img.Height / metafileHeader.DpiY * 100),
                    PixelFormat.Format32bppArgb);

                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.FromArgb(0, Color.White));
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.ScaleTransform(scale, scale);
                    g.DrawImage(img, 0, 0);
                }

                //bitmap.Save(@"c:\temp\test2.png", ImageFormat.Png);
                return bitmap;
            }
        }

        public static Image LoadFromResource(string name, string folder, Assembly assembly)
        {
            string mainAssemblyName = assembly.GetName().Name;
            string resourceName = mainAssemblyName + "." + folder + "." + name;
            System.IO.Stream file = assembly.GetManifestResourceStream(resourceName);
            Image result = Image.FromStream(file);
            return result;
        }

        public static bool IsTheSame(Image imageToCompare, Image imageEthalon)
        {
            //bool sizeTheSame = image1.Size == image2.Size;
            //Debug.Assert(sizeTheSame == false);

            double differences = CompareImages(imageToCompare, imageEthalon, 10);
            return differences < 10.0 ? true : false; //was 2.0 - quite high requirement
        }

        private static double CompareImages(Image image1, Image image2, int Tollerance)
        {
            //https://stackoverflow.com/questions/3384967/how-to-compare-image-objects-with-c-sharp-net
            //with addition of A channel, which is in fact not necessary
            Bitmap Image1 = new Bitmap(image1, new Size(128, 128));
            Bitmap Image2 = new Bitmap(image2, new Size(128, 128));
            int Image1Size = Image1.Width * Image1.Height;
            int Image2Size = Image2.Width * Image2.Height;
            Bitmap Image3;
            if (Image1Size > Image2Size)
            {
                Image1 = new Bitmap(Image1, Image2.Size);
                Image3 = new Bitmap(Image2.Width, Image2.Height);
            }
            else
            {
                Image1 = new Bitmap(Image1, Image2.Size);
                Image3 = new Bitmap(Image2.Width, Image2.Height);
            }
            for (int x = 0; x < Image1.Width; x++)
            {
                for (int y = 0; y < Image1.Height; y++)
                {
                    Color Color1 = Image1.GetPixel(x, y);
                    Color Color2 = Image2.GetPixel(x, y);
                    int a = Color1.A > Color2.A ? Color1.A - Color2.A : Color2.A - Color1.A;
                    int r = Color1.R > Color2.R ? Color1.R - Color2.R : Color2.R - Color1.R;
                    int g = Color1.G > Color2.G ? Color1.G - Color2.G : Color2.G - Color1.G;
                    int b = Color1.B > Color2.B ? Color1.B - Color2.B : Color2.B - Color1.B;
                    Image3.SetPixel(x, y, Color.FromArgb(a, r, g, b));
                }
            }

            //Image1.Save(@"c:\temp\cmp1.png", ImageFormat.Png);
            //Image2.Save(@"c:\temp\cmp2.png", ImageFormat.Png);
            //Image3.Save(@"c:\temp\cmp3.png", ImageFormat.Png);

            int Difference = 0;
            for (int x = 0; x < Image1.Width; x++)
            {
                for (int y = 0; y < Image1.Height; y++)
                {
                    Color Color1 = Image3.GetPixel(x, y);
                    int Media = (Color1.R + Color1.G + Color1.B) / 3;
                    if (Media > Tollerance)
                        Difference++;
                }
            }
            double UsedSize = Image1Size > Image2Size ? Image2Size : Image1Size;
            double result = Difference * 100 / UsedSize;
            return result;
        }

    }
}
