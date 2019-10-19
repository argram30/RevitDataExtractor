#region Namespaces
using System;
using System.Reflection;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
#endregion

namespace RevitDataExtractor
{
    class DataExtractorApp : IExternalApplication
    {
        static void AddRibbonPanel(UIControlledApplication application)
        {
            // Create a custom Ribbon tab
            string tabName = "Wish Box";
            application.CreateRibbonTab(tabName);

            // Add a new ribbon panel
            RibbonPanel tools = application.CreateRibbonPanel(tabName, "Tools");

            // Get dll assembly path
            string dllPath = Assembly.GetExecutingAssembly().Location;

            //create push button for curve total length
            PushButtonData btnOneData = new PushButtonData("cmdRevitFileProcessor", "Revit File" + System.Environment.NewLine + "Processor", dllPath, "RevitDataExtractor.ProcessRevitFiles");
            PushButton btnOne = tools.AddItem(btnOneData) as PushButton;
            btnOne.ToolTip = "Select Multiple lines to obtain Total length";
            BitmapImage btnoneImage = new BitmapImage(new Uri("pack://application:,,,/RevitDataExtractor;component/Resources/revitFileProcessor.png"));
            btnOne.LargeImage = btnoneImage;

        }
        public Result OnStartup(UIControlledApplication a)
        {
            AddRibbonPanel(a);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
