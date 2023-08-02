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

namespace RevitAPITrainingReadToCSV
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            OpenFileDialog openFileDialog=new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "All files(*.*)|*.*";

            string filePath=string.Empty;


            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath=openFileDialog.FileName;
            }

            if(string.IsNullOrEmpty(filePath))
            {
                return Result.Cancelled;
            }

            var lines=File.ReadAllLines(filePath).ToList();

            List<RoomData> roomDataList = new List<RoomData>();
            foreach(var line in lines)
            {
                List<string> values=line.Split(';').ToList();
                roomDataList.Add(new RoomData
                {
                    Name = values[0],
                    Number = values[1]
                });
            }


            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            using(var ts=new Transaction(doc,"Set parameters"))
            {
                ts.Start();
                foreach (RoomData roomData in roomDataList)
                {
                    Room room=rooms.FirstOrDefault(r=>r.Number.Equals(roomData.Number));
                    if (room == null)
                        continue;
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name);
                }
                ts.Commit();
            }

            return Result.Succeeded;
        }
    }
}
