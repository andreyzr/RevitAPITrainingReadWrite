using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPITrainingReadToExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Exel files(*.xlsx) | *.xlsx"
            };

            string filePath=string.Empty;
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }

            if(string.IsNullOrEmpty(filePath))
            {
                return Result.Cancelled;
            }

            var rooms=new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();


            using(FileStream stream =new FileStream(filePath,FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath);
                ISheet sheet = workbook.GetSheetAt(index:0);

                int rowIndex = 0;
                while(sheet.GetRow(rowIndex) != null)
                {
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                        sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        rowIndex++;
                        continue;
                    }

                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    var room=rooms.FirstOrDefault(r => r.Number.Equals(number));

                    if (room == null)
                    {
                        rowIndex++;
                        continue;
                    }
                        
                    using(var ts=new Transaction(doc,"Set parameter"))
                    {
                        ts.Start();
                        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
                        ts.Commit();
                    }
                    rowIndex++;
                }
            }
            return Result.Succeeded;

        }
    }
}
