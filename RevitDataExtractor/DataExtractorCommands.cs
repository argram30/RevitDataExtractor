#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using revExc = Autodesk.Revit.Exceptions;
#endregion

namespace RevitDataExtractor
{
    #region Process Revit files to extract information
    /// <summary>
    /// Process Revit files to extract information
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ProcessRevitFiles : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("B2FD9EF5-3EEC-4E49-8C7C-D8EE50852AAB"));
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Result result;
            Implementations.cmdData = commandData;
            try
            {
                Implementations.OpenWPF();
                result = Result.Succeeded;
            }
            catch (revExc.OperationCanceledException)
            {
                result = Result.Cancelled;
            }
            catch(Exception ex)
            {
                message = ex.Message;
                result = Result.Failed;
            }
            return result;
        }
    }
    #endregion

}
