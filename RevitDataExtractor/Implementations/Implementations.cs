using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using RevApp = Autodesk.Revit.ApplicationServices;

namespace RevitDataExtractor
{
    internal static class Implementations
    {
        #region Properties for Implementations class
        internal static ExternalCommandData cmdData;
        internal static UIApplication UiApp
        {
            get
            {
                return cmdData.Application;
            }
        }

        internal static RevApp.Application CachedApp
        {
            get
            {
                return UiApp.Application;
            }
        }

        public static Document CachedDoc
        {
            get
            {
                return UiApp.ActiveUIDocument.Document;
            }
        }
        #endregion // Properties

        public static RevitFileProcessorUI revitFileProcessorUI { get; set; }

        #region Revit Processing
        internal static void RevitFileProcessor()
        {
            var openOptions = new OpenOptions();
            openOptions.Audit = false;
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach;

            foreach (var file in revitFileProcessorUI.SourceCollection)
            {
                var fullPath = file.FullPath;
                ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(fullPath);
                var openedDoc = CachedApp.OpenDocumentFile(modelPath, openOptions);

                double volColumn = 0.0, volBeam = 0.0, volFloor = 0.0, areaFloor = 0.0, lenBeam = 0.0;

                var fecColumn = new FilteredElementCollector(openedDoc)
                    .OfCategory(BuiltInCategory.OST_Columns)
                    .WhereElementIsNotElementType();

                var fecBeam = new FilteredElementCollector(openedDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType();

                var fecFloor = new FilteredElementCollector(openedDoc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType();

                volColumn = fecColumn.Select(x => x.LookupParameter("Volume").AsDouble()).ToList().Sum();
                volBeam = fecBeam.Select(x => x.LookupParameter("Volume").AsDouble()).ToList().Sum();
                volFloor = fecFloor.Select(x => x.LookupParameter("Volume").AsDouble()).ToList().Sum();
                areaFloor = fecFloor.Select(x => x.LookupParameter("Area").AsDouble()).ToList().Sum();
                lenBeam = fecBeam.Select(x => x.LookupParameter("Length").AsDouble()).ToList().Sum();

                openedDoc.Close(false);
                TaskDialog.Show("Wishbox", $"The Total volume of column is: {volColumn.ToString()} \n The Total volume of beam is: {lenBeam.ToString()}");
            }

        }
        #endregion

        #region Open WPF Window with button click
        internal static void OpenWPF()
        {
            var userControl = new RevitFileProcessorUI();
            var window = new Window();
            window.Content = userControl;
            window.ShowDialog();
            revitFileProcessorUI = userControl;
        }
        #endregion

    }
}
