using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Shapes object
    /// </summary>
    [WrapsCOM("Excel.Shapes", "0002443A-0000-0000-C000-000000000046")]
    public class Shapes : ComWrapper
    {
        private const int POINTS_PER_INCH = 72;

        internal Shapes(object shapesObj)
            : base(shapesObj)
        {
        }

        /// <summary>
        /// Adds a picture to the spreadsheet. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="filename">Filename of the picture</param>
        /// <param name="linkToFile">Should the spreadsheet retain the file's path?</param>
        /// <param name="saveWithDocument">Should the spreadsheet contain the image data? This must be true if linkToFile is false.</param>
        /// <param name="left">Left position, in points</param>
        /// <param name="top">Top position, in points</param>
        /// <param name="width">Width, in points</param>
        /// <param name="height">Height, in points</param>
        /// <returns></returns>
        public Shape AddPicture(string filename, bool linkToFile, bool saveWithDocument, double left, double top, double width, double height)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("Could not find file " + filename);

            return new Shape(InternalObject.GetType().InvokeMember("AddPicture", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject,
                new object[] { filename, linkToFile, saveWithDocument, left, top, width, height }));
        }

        /// <summary>
        /// Adds a picture to the spreadsheet, using the picture's native size. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="filename">Filename of the picture</param>
        /// <param name="linkToFile">Should the spreadsheet retain the file's path?</param>
        /// <param name="saveWithDocument">Should the spreadsheet contain the image data? This must be true if linkToFile is false.</param>
        /// <param name="left">Left position, in points</param>
        /// <param name="top">Top position, in points</param>
        /// <returns></returns>
        public Shape AddPicture(string filename, bool linkToFile, bool saveWithDocument, double left, double top)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("Could not find file " + filename);

            double width, height;
            using (var img = Image.FromFile(filename))
            {
                using (var g = Graphics.FromImage(img))
                {
                    width = img.Width * POINTS_PER_INCH / g.DpiX;
                    height = img.Height * POINTS_PER_INCH / g.DpiY;
                }
            }

            return AddPicture(filename, linkToFile, saveWithDocument, left, top, width, height);
        }

        /// <summary>
        /// Adds a picture to the spreadsheet. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="filename">Filename of the picture</param>
        /// <param name="linkToFile">Should the spreadsheet retain the file's path?</param>
        /// <param name="saveWithDocument">Should the spreadsheet contain the image data? This must be true if linkToFile is false.</param>
        /// <param name="range">The range into which to insert the picture</param>
        /// <param name="fillRange">Should the image fill the range? If not, it will just use the range as a starting point, and will use its native size.</param>
        /// <returns></returns>
        public Shape AddPicture(string filename, bool linkToFile, bool saveWithDocument, Range range, bool fillRange)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("Could not find file " + filename);

            double width, height;
            if (fillRange)
            {
                width = range.Width;
                height = range.Height;
            }
            else
            {
                using (var img = Image.FromFile(filename))
                {
                    using (var g = Graphics.FromImage(img))
                    {
                        width = img.Width * POINTS_PER_INCH / g.DpiX;
                        height = img.Height * POINTS_PER_INCH / g.DpiY;
                    }
                }
            }

            return AddPicture(filename, linkToFile, saveWithDocument, range.Left, range.Top, width, height);
        }

        /// <summary>
        /// Adds a picture to the spreadsheet. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="filename">Filename of the picture</param>
        /// <param name="linkToFile">Should the spreadsheet retain the file's path?</param>
        /// <param name="saveWithDocument">Should the spreadsheet contain the image data? This must be true if linkToFile is false.</param>
        /// <param name="origin">The range at which to put the top left corner of the image.</param>
        /// <param name="width">Width, in points</param>
        /// <param name="height">Height, in points</param>
        /// <returns></returns>
        public Shape AddPicture(string filename, bool linkToFile, bool saveWithDocument, Range origin, double width, double height)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("Could not find file " + filename);

            return AddPicture(filename, linkToFile, saveWithDocument, origin.Left, origin.Top, width, height);
        }

        /// <summary>
        /// Adds a picture to the spreadsheet. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="img">The image to add</param>
        /// <param name="range">The range into which to insert the picture</param>
        /// <param name="fillRange">Should the image fill the range? If not, it will just use the range as a starting point, and will use its native size.</param>
        /// <returns></returns>
        public Shape AddPicture(Image img, Range range, bool fillRange)
        {

            string tempFile = Path.GetTempFileName();
            File.Move(tempFile, tempFile = tempFile + ".bmp");

            img.Save(tempFile, System.Drawing.Imaging.ImageFormat.Bmp);

            Shape shape = AddPicture(tempFile, false, true, range, fillRange);

            File.Delete(tempFile);

            return shape;
        }

        /// <summary>
        /// Adds a picture to the spreadsheet. The returned shape object must be manually disposed.
        /// </summary>
        /// <param name="img">The image to add</param>
        /// <param name="origin">The range at which to put the top left corner of the image.</param>
        /// <param name="width">Width, in points.</param>
        /// <param name="height">Height, in points.</param>
        /// <returns></returns>
        public Shape AddPicture(Image img, Range origin, double width, double height)
        {
            string tempFile = Path.GetTempFileName();
            File.Move(tempFile, tempFile = tempFile + ".bmp");

            img.Save(tempFile, System.Drawing.Imaging.ImageFormat.Bmp);

            Shape shape = AddPicture(tempFile, false, true, origin, width, height);

            File.Delete(tempFile);

            return shape;
        }
    }
}
