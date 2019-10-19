using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;

namespace RevitDataExtractor
{
    #region DetailLine Filter using BIC
    internal class DetailLineFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    #endregion

    #region Generic Filter with string as Constructor
    internal class GenericFilter : ISelectionFilter
    {
        static string CategoryName = "";

        public GenericFilter(string name)
        {
            CategoryName = name;
        }

        public bool AllowElement(Element elem)
        {
            return elem.Category.Name == CategoryName;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    #endregion
}
