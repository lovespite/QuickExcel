using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System.Windows.Media.Imaging
{
    public class PortableFontDesc
    {
        public readonly string FontName;
        public readonly int EmSize;
        public readonly bool IsBold;
        public readonly bool IsItalic;
        public readonly bool IsClearType;
        /// <summary>Device pixels per DIP horizontally used to rasterize glyphs.</summary>
        public readonly double ScaleX;
        /// <summary>Device pixels per DIP vertically used to rasterize glyphs.</summary>
        public readonly double ScaleY;

        public PortableFontDesc(string name = "Arial", int emsize = 12, bool isbold = false, bool isitalic = false, bool cleartype = false,
            double scaleX = 0, double scaleY = 0)
        {
            FontName = name;
            EmSize = emsize;
            IsBold = isbold;
            IsItalic = isitalic;
            IsClearType = cleartype;
            // A non-positive scale means the caller still relies on the legacy
            // window DPI; hosts that render under a layout zoom pass the actual
            // device scale so glyphs keep one-to-one device pixels.
            ScaleX = scaleX > 0 ? scaleX : DpiDetector.DpiXKoef;
            ScaleY = scaleY > 0 ? scaleY : DpiDetector.DpiYKoef;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return FontName.GetHashCode() ^ EmSize.GetHashCode() ^ IsBold.GetHashCode() ^ IsItalic.GetHashCode() ^
                    IsClearType.GetHashCode() ^ ScaleX.GetHashCode() ^ ScaleY.GetHashCode();
            }
        }

        public override bool Equals(object obj)
        {
            var other = obj as PortableFontDesc;
            if (other == null) return false;
            return FontName == other.FontName && EmSize == other.EmSize && IsBold == other.IsBold && IsItalic == other.IsItalic &&
                IsClearType == other.IsClearType && ScaleX.Equals(other.ScaleX) && ScaleY.Equals(other.ScaleY);
        }
    }
}
