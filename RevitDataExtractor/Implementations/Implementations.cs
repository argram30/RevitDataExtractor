using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;

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

        public static Application CachedApp
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

        internal static string pathName { get; set; }

        #region Revit Processing
        internal static void RevitFileProcessor()
        {
            DirectoryInfo dirInfo = new DirectoryInfo(pathName);
            var files = dirInfo.GetFiles("*.rvt");
            var openOptions = new OpenOptions();
            openOptions.Audit = false;
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach;

            foreach (var file in files)
            {
                ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(file.FullName);
                var openedDoc = CachedApp.OpenDocumentFile(modelPath, openOptions);

                double volColumn, volBeam, volFloor, areaFloor, lenBeam;

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
            }



            // Do things here
            TaskDialog.Show("Wishbox", "Processing....");
        }
        #endregion

        #region Open WPF Window with button click
        internal static void OpenWPF()
        {

        }
        #endregion
    }
}
