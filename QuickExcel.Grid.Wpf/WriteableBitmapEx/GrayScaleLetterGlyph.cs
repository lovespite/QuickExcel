using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace System.Windows.Media.Imaging
{
    public class GrayScaleLetterGlyph
    {
        public struct Item
        {
            public short X;
            public short Y;
            public int Alpha;
        }

        public char Ch;
        public int Width;
        public int Height;

        public Item[] Items;


        public static GrayScaleLetterGlyph CreateSpaceGlyph(GlyphTypeface glyphTypeface, double size, double scaleX, double scaleY)
        {
            int spaceWidth = (int)Math.Ceiling(glyphTypeface.AdvanceWidths[glyphTypeface.CharacterToGlyphMap[' ']] * size * scaleX);
            return new GrayScaleLetterGlyph
            {
                Ch = ' ',
                Height = (int)Math.Ceiling(glyphTypeface.Height * size * scaleY),
                Width = spaceWidth,
            };
        }

        public static unsafe GrayScaleLetterGlyph CreateGlyph(Typeface typeface, GlyphTypeface glyphTypeface, double size, char ch,
            double scaleX = 0, double scaleY = 0)
        {
            if (scaleX <= 0) scaleX = DpiDetector.DpiXKoef;
            if (scaleY <= 0) scaleY = DpiDetector.DpiYKoef;
            if (ch == ' ') return CreateSpaceGlyph(glyphTypeface, size, scaleX, scaleY);

            FormattedText text = new FormattedText("" + ch,
                                                   CultureInfo.InvariantCulture,
                                                   FlowDirection.LeftToRight,
                                                   typeface,
                                                   size,
                                                   Brushes.White, scaleY);

            int width = (int)Math.Ceiling(scaleX * text.Width);
            int height = (int)Math.Ceiling(scaleY * text.Height);
            if (width == 0 || height == 0) return null;

            DrawingVisual drawingVisual = new DrawingVisual();
            DrawingContext drawingContext = drawingVisual.RenderOpen();
            drawingContext.DrawRectangle(Brushes.Black, new Pen(), new Rect(0, 0, width, height));
            drawingContext.DrawText(text, new Point(0, 0));
            drawingContext.Close();

            RenderTargetBitmap bmp = new RenderTargetBitmap(width, height, scaleX * 96, scaleY * 96, PixelFormats.Pbgra32);
            bmp.Render(drawingVisual);

            var res = new List<Item>();

            var pixbmp = new WriteableBitmap(bmp);
            using (var ctx = new BitmapContext(pixbmp))
            {
                var pixels = ctx.Pixels;

                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        int color = pixels[y*width + x];

                        byte r, g, b;
                        double avg;

                        r = (byte) ((color >> 16) & 0xFF);
                        g = (byte) ((color >> 8) & 0xFF);
                        b = (byte) ((color) & 0xFF);

                        avg = 0.299*r + 0.587*g + 0.114*b;

                        if (avg >= 1)
                        {
                            res.Add(new Item
                                {
                                    X = (short) x,
                                    Y = (short) y,
                                    Alpha = (int) Math.Round(avg/255.0*0x1000),
                                });
                        }
                    }
                }
            }

            return new GrayScaleLetterGlyph
                {
                    Width = width,
                    Height = height,
                    Ch = ch,
                    Items = res.ToArray(),
                };
        }
    }
}
