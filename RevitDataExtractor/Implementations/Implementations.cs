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

        #region Curve Total Length
        internal static void RevitFileProcessor()
        {
            // Do things here
            TaskDialog.Show("Wishbox", "Processing....");
        }
        #endregion

    }
}
