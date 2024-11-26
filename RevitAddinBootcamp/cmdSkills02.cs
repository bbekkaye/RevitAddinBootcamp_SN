using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.Exceptions;
using System.Windows.Controls;





namespace RevitAddinBootcamp
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSkills02 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            // 1a. pick a single element
            Reference pickRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select Element");
            Element pickElement = doc.GetElement(pickRef);

            ////////////////////////////////////////////////////////////////////////////////////////////////////

            // 1b. pick multiple elements
            List<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select Elements").ToList();

            TaskDialog.Show("Revit Selection", $"{pickList.Count} elements selected.");

            ////////////////////////////////////////////////////////////////////////////////////////////////////

            // 2.filter selected elements for lines

            List<CurveElement> allCurves = new List<CurveElement>();
             foreach (Element elem in pickList)
             {
                 if (elem is CurveElement)
                 {
                     allCurves.Add(elem as CurveElement);
                }
             }

             ///////////////////////////////////////////////////////////////////////////////////////////////////////

            // 2b.filter selected elements for model curves

            List < CurveElement > modelCurves = new List<CurveElement>();
            foreach (Element elem2 in pickList)
            {
                if (elem2 is CurveElement)
                {
                    // CurveElement curveElem = elem2 as CurveElement;
                    CurveElement curveElem = (CurveElement)elem2;

                    if (curveElem.CurveElementType == CurveElementType.ModelCurve)
                    {
                        modelCurves.Add(curveElem);
                    }
                }
            }

            // 3.curve data
             foreach (CurveElement currentCurve in modelCurves)
            {
                Curve curve = currentCurve.GeometryCurve;
                XYZ startPoint = curve.GetEndPoint(0);
                XYZ endPoint = curve.GetEndPoint(1);

                GraphicsStyle curStyle = currentCurve.LineStyle as GraphicsStyle;

                Debug.Print(curStyle.Name);

            }

            // 5.Create a Transaction with using statement
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Wall");

                // 4.create a wall
                Level newLevel = Level.Create(doc, 20);

                CurveElement curveElem = modelCurves[0];
                Curve curCurve = curveElem.GeometryCurve;
                Curve curCurve2 = modelCurves[1].GeometryCurve;

                Wall newWall = Wall.Create(doc, curCurve, newLevel.Id, false);

                //////////////////////////////////////////////////////////////////////////////////////////

                // 4b. create a wall with wall type
                FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
                wallTypes.OfCategory(BuiltInCategory.OST_Walls);
                wallTypes.WhereElementIsElementType();

                Wall newWall2 = Wall.Create(doc, curCurve2, wallTypes.FirstElementId(), newLevel.Id, 20, 0, false, false);

                ///////////////////////////////////////////////////////////////////////


                // 6. get system types
                FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
                systemCollector.OfClass(typeof(MEPSystemType));

                // 7. get duct system type
                MEPSystemType ductSystem = GetSystemTypeByName(doc, "Supply Air");

                
                // 8.Get duct type
                FilteredElementCollector ductCollector = new FilteredElementCollector(doc);
                ductCollector.OfClass(typeof(DuctType));


                // 9. create duct
                Curve ductCurve = modelCurves[2].GeometryCurve;
                XYZ ductStart = ductCurve.GetEndPoint(0);
                XYZ ductEnd = ductCurve.GetEndPoint(1);

                Duct newDuct = Duct.Create(doc, ductSystem.Id, ductCollector.FirstElementId(),
                    newLevel.Id, ductStart, ductEnd);

                /////////////////////////////////////////////////////////////////////////////////////

                // 10. get pipe system type
                MEPSystemType pipeSystem = GetSystemTypeByName(doc, "Domestic Hot Water");

                // 11. get pipe type
                FilteredElementCollector pipeCollector = new FilteredElementCollector(doc);
                pipeCollector.OfClass(typeof(PipeType));

                // 12. create pipe
                Curve pipeCurve = modelCurves[3].GeometryCurve;
                XYZ pipeStart = pipeCurve.GetEndPoint(0);
                XYZ pipeEnd = pipeCurve.GetEndPoint(1);
                
                Pipe newPipe = Pipe.Create(doc, pipeSystem.Id, pipeCollector.FirstElementId(),
                    newLevel.Id, pipeStart, pipeEnd);


                t.Commit();

            }


            return Result.Succeeded;
        }

        internal MEPSystemType GetSystemTypeByName(Document doc, string name)
        {
            // get all system types
            FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
            systemCollector.OfClass(typeof(MEPSystemType));

            // get duct system type by name

            MEPSystemType systemType = systemCollector.Cast<MEPSystemType>().FirstOrDefault(x => x.Name == name);

            return systemType;
        }
    }

}
