using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;

namespace ExcelDrawing
{
    class Picture
    {
        BitmapData bData;

        public string Path { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public Image Img { get; set; }
        public Bitmap Picture1 { get; set; }

        public Picture(string path)
        {
            Path = path;
            Img = Image.FromFile(path);
            Picture1 = (Bitmap)Img;
            Width = Picture1.Size.Width;
            Height = Picture1.Size.Height;
            //bData = Picture1.LockBits(new Rectangle(0, 0, Img.Width, Img.Height), ImageLockMode.ReadWrite, Img.PixelFormat);
        }

        public Picture(string path, int height, int width)
        {
            Img = Image.FromFile(path);
            Picture1 = new Bitmap(Image.FromFile(path), new Size(width, height));
            Path = path;
            Width = width;
            Height = height;
            //bData = Picture1.LockBits(new Rectangle(0, 0, Img.Width, Img.Height), ImageLockMode.ReadWrite, Img.PixelFormat);
        }

        public Color GetColor(int x, int y)
        {
            Color color = Picture1.GetPixel(x, y);

            return color;
        }
    }
}
