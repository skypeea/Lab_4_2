using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_4_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            var pipeList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            string exelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");
            using (FileStream stream = new FileStream(exelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");
                int rowIndex = 0;
                sheet.SetCellValue(rowIndex, columnIndex: 0, "Name");
                sheet.SetCellValue(rowIndex, columnIndex: 1, "Length, m");
                sheet.SetCellValue(rowIndex, columnIndex: 2, "Outer diameter, mm");
                sheet.SetCellValue(rowIndex, columnIndex: 3, "Inner diameter, mm");
                rowIndex++;
                foreach (var pipe in pipeList)
                {
                    
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 1, ToMeters(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()));
                    sheet.SetCellValue(rowIndex, columnIndex: 2, ToMillimeters(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble()));
                    sheet.SetCellValue(rowIndex, columnIndex: 3, ToMillimeters(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble()));
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();

            }
            
            
            return Result.Succeeded;

        }


         double ToMeters(double c)
        {
           return UnitUtils.ConvertFromInternalUnits(c, UnitTypeId.Meters);
        }
         double ToMillimeters(double c)
        {
            return UnitUtils.ConvertFromInternalUnits(c, UnitTypeId.Millimeters);
        }
    }
}
