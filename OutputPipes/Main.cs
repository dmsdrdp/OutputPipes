using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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

//Задание 4.2 Вывод значений труб

namespace OutputPipes
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //string roomInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string exelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");
            using (FileStream stream = new FileStream(exelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    Parameter outerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);          //параметры
                    Parameter innerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);     
                    Parameter length = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);


                    double outerDiameterConverted = UnitUtils.ConvertFromInternalUnits(outerDiameter.AsDouble(), UnitTypeId.Millimeters); //конвертированные в миллиметры параметры
                    double innerDiameterConverted = UnitUtils.ConvertFromInternalUnits(innerDiameter.AsDouble(), UnitTypeId.Millimeters);
                    double lengthConverted = UnitUtils.ConvertFromInternalUnits(length.AsDouble(), UnitTypeId.Millimeters);

                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.Name);                         //вывод параметров в exel
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outerDiameterConverted);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, innerDiameterConverted);       
                    sheet.SetCellValue(rowIndex, columnIndex: 3, lengthConverted);
                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }
            System.Diagnostics.Process.Start(exelPath);

            return Result.Succeeded;

        }
    }
}