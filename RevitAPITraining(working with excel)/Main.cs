using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPITraining_working_with_excel_
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            foreach (Room room in rooms)
            {
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                roomInfo += $"{room.Name}\t{room.Number}\t{room.Area}{Environment.NewLine}";
            }

            //string desktopPath=Environment.GetFolderPath(Environment.SpecialFolder.Desktop);  Сохранение на рабочем столе
            //string csvPath = Path.Combine(desktopPath, "roomInfo.csv");

            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files(*.*)|*.*",
                FileName = "roomInfo.csv",
                DefaultExt = ".csv"
            };

            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
            {
                return Result.Cancelled;
            }

            File.WriteAllText(selectedFilePath, roomInfo, Encoding.Unicode);

            return Result.Succeeded;

        }
    }
}
