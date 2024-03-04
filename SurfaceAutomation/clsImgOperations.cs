using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SurfaceAutomation
{
    public class clsImgOperations
    {
        public void ScreenCapture(string path,int height, int width)
        {
            Bitmap bmpScreenshot = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            Graphics g = Graphics.FromImage(bmpScreenshot);
            g.CopyFromScreen(42, 36, width, height, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
            bmpScreenshot.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
            bmpScreenshot.Dispose();
        }
        public void CropImage(int x, int y, int width, int height, string inPath, string outPath, string errorMsg)
        {
            errorMsg = string.Empty;
            try
            {
                Bitmap croppedImage;
                using(var orginalImage = new Bitmap(inPath))
                {
                    System.Drawing.Rectangle crop = new System.Drawing.Rectangle(x, y, width, height);
                    croppedImage = orginalImage.Clone(crop, orginalImage.PixelFormat);
                }
                croppedImage.Save(outPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                croppedImage.Dispose();
            }
            catch(Exception ex)
            {
                errorMsg = ex.Message;
            }
        }
    }

    class ColorMatrix
    {
        public ColorMatrix()
        {

        }
        public float[][] Matrix { get; set; }

        public Bitmap Apply(Bitmap OrginalImage)
        {
            Bitmap NewBitmap = new Bitmap(OrginalImage.Width, OrginalImage.Height);
            using (Graphics NewGraphics = Graphics.FromImage(NewBitmap))
            {
                System.Drawing.Imaging.ColorMatrix NewColorMatrix = new System.Drawing.Imaging.ColorMatrix(Matrix);
                using(ImageAttributes Attributes = new ImageAttributes())
                {
                    Attributes.SetColorMatrix(NewColorMatrix);
                    NewGraphics.DrawImage(OrginalImage, new System.Drawing.Rectangle(0,0,OrginalImage.Width,OrginalImage.Height), 0, 0, OrginalImage.Width, OrginalImage.Height, GraphicsUnit.Pixel, Attributes);
                }
            }
            return NewBitmap;
        }
    }
}
