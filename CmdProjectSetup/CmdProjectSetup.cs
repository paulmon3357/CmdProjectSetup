#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddinBootcamp
{
    [Transaction(TransactionMode.Manual)]
    public class CmdProjectSetup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //TaskDialog.Show("Test", "CmdProjectSetup");
            string filepath = @"M:\OneDrive - Mon & Associates Inc\Library\REIVT\REVIT ADDIN BOOTCAMP\Download\RAB_Session_02_Challenge_Levels.csv";
            string fileText = System.IO.File.ReadAllText(filepath);
            string[] fileArray = System.IO.File.ReadAllLines(filepath);
            int LineNumber = 1;
            //TaskDialog.Show("Test",fileText);

            Transaction t = new Transaction(doc);
            t.Start("Create level");

            foreach (string rowString in fileArray)
            {
                string[] cellString = rowString.Split(',');

                string levelName = cellString[0];
                string levelHeightString = cellString[1];
                double levelHeightNumber = 0.00;
                string LineNumberString = LineNumber.ToString();

                try
                {
                    // create level
                    levelHeightNumber = Convert.ToDouble(levelHeightString);
                    Level myLevel = Level.Create(doc, levelHeightNumber);
                    myLevel.Name = levelName;

                }
                catch
                {
                    TaskDialog.Show("Exception", "Error on data file line number" + LineNumberString +" !\n" +"Cannot convert \"" + levelHeightString + "\" level height to double!");
                }





                //FilteredElementCollector collector = new FilteredElementCollector(doc);
                //collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                //ElementId tblockId = collector.FirstElementId();

                LineNumber++;
               


            }


                t.Commit();
                t.Dispose();

            string filepath2 = @"m:\onedrive - mon & associates inc\library\reivt\revit addin bootcamp\download\rab_session_02_challenge_sheets.csv";
            string[] fileArray2 = System.IO.File.ReadAllLines(filepath2);
            LineNumber = 1;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = collector.FirstElementId();
            


            //--------------Start Transtraction----------------------------------------------//


            Transaction t2 = new Transaction(doc);
            t2.Start("Create Sheets");

            foreach (string rowString2 in fileArray2)
            {
                string[] cellString2 = rowString2.Split(',');
                string sheetNo = cellString2[0];
                string SheetName = cellString2[1];

                // create sheet
                if (LineNumber!= 1)
                {
                    ViewSheet mySheet = ViewSheet.Create(doc, tblockId);
                    mySheet.Name = SheetName;
                    mySheet.SheetNumber = sheetNo;
                }

                string LineNumberString = LineNumber.ToString();
                Debug.Print(LineNumberString);

                LineNumber++;
                
                
            }

            t2.Commit();
            t2.Dispose();
            //------------------------------end transaction------------------------------------//


            return Result.Succeeded;
            
        }


    }
}
