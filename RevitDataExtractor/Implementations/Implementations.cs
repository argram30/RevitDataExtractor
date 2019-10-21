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
using System.Collections.ObjectModel;


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


        #region Revit Processing
        internal static void RevitFileProcessor(ObservableCollection<RevitFile> filenames)
        {
            var openOptions = new OpenOptions();
            openOptions.Audit = false;
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach;
            var num = 2;
            var strList = new List<string> { "File Name", "Volume of Column", "Volume of Beam", "Volume of Floor", "Area of Floor", "Length of Beam" };

            Excel excel = new Excel();
            excel.initNew();
            excel.open(@"C:\Users\Raja\Desktop\hackathon test files\datawriter.xlsx");
            excel.writeRow("Tabelle1", 1, strList);

            foreach (var file in filenames)
            {
                var fullPath = file.FullPath;
                var excelFile = file.Directory;
                ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(fullPath);
                var openedDoc = CachedApp.OpenDocumentFile(modelPath, openOptions);

                double volColumn = 0.0, volBeam = 0.0, volFloor = 0.0, areaFloor = 0.0, lenBeam = 0.0;

                var fecColumn = new FilteredElementCollector(openedDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
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

                var valList = new List<string> { file.Filename, volColumn.ToString(), volBeam.ToString(), volFloor.ToString(), areaFloor.ToString(), lenBeam.ToString() };
                
                excel.writeRow("Tabelle1", num, valList);

                num += 1;

                openedDoc.Close(false);
                //TaskDialog.Show("Wishbox", $"The Total volume of column is: {volColumn.ToString()} \n The Total volume of beam is: {lenBeam.ToString()}");
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
            //revitFileProcessorUI = userControl;
        }
        #endregion

    }
}
