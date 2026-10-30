using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunExtendedStyleSmokeAsync(string folder, Action<bool, string> check)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Style layout", cells: [
            new(new(0,0),new(CellContent.FromLiteral(CellValue.FromText("Rotate ABC"),1))),
            new(new(1,0),new(CellContent.FromLiteral(CellValue.FromText("Indented"),2))),
            new(new(2,0),new(CellContent.FromLiteral(CellValue.FromText("Long text should shrink to the cell width"),3)))],
            columns:[new(0,new(24))],rows:[new(0,new(100))],defaultRowHeight:40);
        var snapshot=new WorkbookSnapshot(Guid.NewGuid(),0,ExcelDateSystem.Windows1900,[sheet],styles:[new(),
            new(){FontName="Arial",EastAsianFontName="Microsoft YaHei",FontSize=16,TextRotation=45},
            new(){FontName="Arial",FontSize=16,Indent=2},new(){FontName="Arial",FontSize=16,ShrinkToFit=true}]);
        InstallDocument(new WorkbookDocument(snapshot,_services),null,"扩展样式验收"); GridSurface.Select(new(8,5));
        await Dispatcher.InvokeAsync(UpdateLayout,DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control);
        var bitmap=new RenderTargetBitmap((int)Math.Ceiling(_grid.Artwork.ActualWidth),(int)Math.Ceiling(_grid.Artwork.ActualHeight),96,96,PixelFormats.Pbgra32);
        bitmap.Render(_grid.Artwork);var pixels=new byte[bitmap.PixelWidth*bitmap.PixelHeight*4];bitmap.CopyPixels(pixels,bitmap.PixelWidth*4,0);
        Rect Ink(int row)
        {
            var cell=_grid.Control.GetModelCellBounds(row,0)!.Value;var result=Rect.Empty;
            for(var y=(int)cell.Top+2;y<cell.Bottom-2;y++) for(var x=(int)cell.Left+2;x<cell.Right-2;x++)
            {
                var index=(y*bitmap.PixelWidth+x)*4;
                if(pixels[index]<100 && pixels[index+1]<100 && pixels[index+2]<100) result.Union(new Point(x-cell.Left,y-cell.Top));
            }
            return result;
        }
        check(Ink(0).Height>40,"Rotated text paints an angled line with a taller glyph bounding box");
        check(Ink(1).Left>20,"Indent moves actual glyphs away from the cell edge");
        check(Ink(2).Width>100 && Ink(2).Height<18,"Shrink-to-fit renders the full long line at a smaller glyph size");
        var fitted=FastGridArtwork.CreateText(sheet.GetCell(2,0)!.Value.TextValue,snapshot.Styles[3],VisualTreeHelper.GetDpi(_grid.Artwork).PixelsPerDip);
        var available=_grid.Control.GetModelCellBounds(2,0)!.Value.Width-4; FastGridArtwork.FitText(fitted,snapshot.Styles[3],available);
        check(fitted.WidthIncludingTrailingWhitespace<=available,"Shrink-to-fit remeasures hinted glyphs so the complete line fits without ellipsis");
        check(_grid.MeasureAutoFit(SheetAxis.Row,0)>50,"Auto-fit row height includes rotated text bounds");
        CaptureGridEvidence(Path.Combine(folder,"extended-styles.png"));
    }
}
