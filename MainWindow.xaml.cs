using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptApplication = Microsoft.Office.Interop.PowerPoint.Application;

namespace PowerPointSlideExportDemo
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private readonly string _pptxFilePath;
		private readonly string _pngFileBasePath;

		private const int _PngWidth = 1920;
		private const int _PngHeight = 1080;

		public MainWindow()
		{
			InitializeComponent();

			_pptxFilePath = Path.Combine(AppContext.BaseDirectory, "Example.pptx");
			_pngFileBasePath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

			UpdateDpiInfoText();
			DpiChanged += (sender, args) => UpdateDpiInfoText(args.OldDpi,args.NewDpi);
		}

		private void ExportSlideWithBackground_Click(object sender, RoutedEventArgs e)
		{
			var pngFilePath = Path.Combine(_pngFileBasePath, "SlideWithBackground.png");

			var powerPoint = new PptApplication();
			var presentation = powerPoint.Presentations.Open(_pptxFilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
			var slide = presentation.Slides[1]; // one-based!
			slide.Export(pngFilePath, "PNG", _PngWidth, _PngHeight);
		}

		private void ExportSlideWithoutBackground_Click(object sender, RoutedEventArgs e)
		{
			var pngFilePath = Path.Combine(_pngFileBasePath, "SlideWithoutBackground.png");

			var powerPoint = new PptApplication();
			var presentation = powerPoint.Presentations.Open(_pptxFilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
			var slide = presentation.Slides[1]; // one-based!
			var shapes = slide.Shapes;

			var pageSetup = presentation.PageSetup;
			var rectangle = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, pageSetup.SlideWidth, pageSetup.SlideHeight);
			rectangle.Fill.Visible = MsoTriState.msoFalse;
			rectangle.Line.Visible = MsoTriState.msoFalse;

			var range = shapes.Range();
			range.Export(
				pngFilePath,
				PpShapeFormat.ppShapeFormatPNG,
				(int)(_PngWidth * 72 / 96),
				(int)(_PngHeight * 72 / 96),
				PpExportMode.ppScaleXY);
		}

		private void UpdateDpiInfoText()
		{
			var dpiScale = VisualTreeHelper.GetDpi(this);
			DpiInfoText.Text = $"Pixels per inch at startup: X={dpiScale.PixelsPerInchX}, Y={dpiScale.PixelsPerInchY}";
		}


		private void UpdateDpiInfoText(DpiScale oldDpi, DpiScale newDpi)
		{
			var dpiScale = VisualTreeHelper.GetDpi(this);
			DpiInfoText.Text = $"Pixels per inch changed from (X={oldDpi.PixelsPerInchX}, Y={oldDpi.PixelsPerInchY}) to (X={newDpi.PixelsPerInchX}, Y={newDpi.PixelsPerInchY})";
		}
	}
}
