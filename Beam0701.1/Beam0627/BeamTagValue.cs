using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;
using Document = Autodesk.Revit.DB.Document;
using Autodesk.Revit.ApplicationServices;
using System.Runtime.InteropServices;
using System.Diagnostics;
using GeoElement = Autodesk.Revit.DB.GeometryElement;
using FaceUtils = Autodesk.Revit.DB.HostObjectUtils;
using System.Data;

namespace Beam0627
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    public class BeamTagValue : IExternalCommand
    {
        UIApplication app;
        Document doc;
        String info;

        public Result Execute(ExternalCommandData commandData, ref string messages, ElementSet elements)
        {
            app = commandData.Application;
            doc = app.ActiveUIDocument.Document;
            string prompt = null;

            /************************在專案新增共用參數BTV************************/
            StringBuilder levelOrderInformation2 = new StringBuilder();
            Category category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFraming);
            CategorySet categorySet = app.Application.Create.NewCategorySet();
            categorySet.Insert(category);
            string originalFile = app.Application.SharedParametersFilename;
            string tempFile = @"C:\BeamTagValue.txt";
            //string tempFile = "Data source=myData.db";
            
            app.Application.SharedParametersFilename = tempFile;
            DefinitionFile shareParameterFile = app.Application.OpenSharedParameterFile();
            try
            {
                foreach (DefinitionGroup dg in shareParameterFile.Groups)
                {
                    if (dg.Name == "GROUP1")
                    {
                        ExternalDefinition externalDefinition = dg.Definitions.get_Item("BTV") as ExternalDefinition;
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Add Shared Parameters");
                            InstanceBinding newIB = app.Application.Create.NewInstanceBinding(categorySet);
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_TEXT);
                            t.Commit();
                        }
                    }
                }
            }
            catch { }
            finally
            {
                app.Application.SharedParametersFilename = originalFile;
            }
            try
            {
                /************************找整個模型所有標高************************/
                StringBuilder levelInformation = new StringBuilder();
                int levelNumber = 0;
                // Get the handle of current document.
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                FilteredElementCollector collectorLevel = new FilteredElementCollector(document);
                ICollection<Element> collectionLevel = collectorLevel.OfClass(typeof(Level)).ToElements();
                string[,] LevelInfo = new string[10, 4];
                foreach (Element e in collectionLevel)
                {
                    Level level = e as Level;
                    if (level != null)
                    {
                        LevelInfo[levelNumber, 0] = level.Name.ToString();
                        LevelInfo[levelNumber, 1] = level.Id.ToString();
                        LevelInfo[levelNumber, 2] = level.Elevation.ToString();
                        levelNumber++;
                    }
                }
                levelInformation.Append("本專案總共有 " + levelNumber + " 個levels。");
                TaskDialog.Show("Revit", levelInformation.ToString());

                /************************排序模型所有標高(由上至下)************************/
                string[,] temp = new string[10, 3];
                for (int i = 0; i < 10; i++)
                {
                    for (int j = i + 1 ; j < 10; j++)
                    {
                        if (Convert.ToDouble(LevelInfo[i, 2]) < Convert.ToDouble(LevelInfo[j, 2]))
                        {
                            temp[0, 2] = LevelInfo[i, 2];
                            LevelInfo[i, 2] = LevelInfo[j, 2];
                            LevelInfo[j, 2] = temp[0, 2];

                            temp[0, 0] = LevelInfo[i, 0];
                            LevelInfo[i, 0] = LevelInfo[j, 0];
                            LevelInfo[j, 0] = temp[0, 0];

                            temp[0, 1] = LevelInfo[i, 1];
                            LevelInfo[i, 1] = LevelInfo[j, 1];
                            LevelInfo[j, 1] = temp[0, 1];
                        }
                    }
                }
                StringBuilder levelOrderInformation = new StringBuilder();
                for (int i = 0; i < 10; i++)
                {
                    int j = 0;
                    LevelInfo[i, j + 3] = i.ToString();
                    levelOrderInformation.Append("\nLevel名稱: " + LevelInfo[i, j]);
                    levelOrderInformation.Append("\nLevelID: " + LevelInfo[i, j + 1]);
                    levelOrderInformation.Append("\nLevel高度: " + Math.Round(Convert.ToDouble(LevelInfo[i, j + 2]), 2));
                    //levelOrderInformation.Append("\nLevelKey: " + LevelInfo[i, j + 3]);
                    levelOrderInformation.Append(" 公尺\n");
                }
                TaskDialog.Show("Revit", levelOrderInformation.ToString());

                

                /************************找整個模型所有的樑************************/
                ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(FamilyInstance));
                ElementCategoryFilter beamsCategoryfilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                LogicalAndFilter beamInstancesFilter = new LogicalAndFilter(familyInstanceFilter, beamsCategoryfilter);
                FilteredElementCollector collector = new FilteredElementCollector(document);
                ICollection<ElementId> beams = collector.WherePasses(beamInstancesFilter).ToElementIds();
                IList<Element> beam = collector.WherePasses(beamInstancesFilter).ToElements();
                foreach (Element element in beam)
                {
                    Parameter p = element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    Parameter btv_param = element.LookupParameter("BTV");
                    //Parameter offset = element.LookupParameter("o");
                    //double height = offset.AsDouble();
                    //Parameter offset = element.get_Parameter(BuiltInParameter.height);



                    /*
                    FilteredElementCollector a = new FilteredElementCollector(document)
.                   OfClass(typeof(SharedParameterElement));
                    s = new StringBuilder();
                    foreach (var r in a)
                    {
                        SharedParameterElement shElem = r as SharedParameterElement;
                        InternalDefinition def = shElem.GetDefinition();
                        if(def.Name == "BTV")
                        {
                            s.AppendLine(string.Format("{0}, {1} [{2}]", def.Name, def.ParameterType, shElem.GuidValue));
                        }
                    }
                    */





                    using (Autodesk.Revit.DB.Transaction trans = new Autodesk.Revit.DB.Transaction(doc, "RESI"))
                    {
                        trans.Start();
                        string RefLevelID = GetParameterValue(p);
                        double BeamTagValue = 0;
                        //string Offset = offset.AsValueString();
                        try
                        {
                            for (int i = 0; i < 10; i++)
                            {
                                if (RefLevelID == LevelInfo[i, 1] && i != 10)
                                {
                                    BeamTagValue = Convert.ToDouble(LevelInfo[i, 2]) - Convert.ToDouble(LevelInfo[i + 1, 2]);
                                    BeamTagValue = Math.Round(BeamTagValue*30.48, 2);
                                    btv_param.Set(BeamTagValue/30.48);
                                    prompt += "\n\t" + element.Name + " <" + GetParameterValue(p) + ">" + " <" + BeamTagValue.ToString() + ">";
                                    //prompt += "\n\t" + element.Name + " <" + GetParameterValue(p) + ">" + " <" + BeamTagValue.ToString() + ">" + " <" + offset + ">";
                                }
                                else if (RefLevelID == LevelInfo[i, 1] && i == 9)
                                {
                                    BeamTagValue = Convert.ToDouble(LevelInfo[i, 2]);
                                    BeamTagValue = Math.Round(BeamTagValue* 30.48, 2);
                                    btv_param.Set(BeamTagValue/30.48);
                                    prompt += "\n\t" + element.Name + " <" + GetParameterValue(p) + ">" + " <" + BeamTagValue.ToString() + ">";
                                    //prompt += "\n\t" + element.Name + " <" + GetParameterValue(p) + ">" + " <" + BeamTagValue.ToString() + ">" + " <" + offset + ">";
                                }
                            }
                        }
                        catch
                        {
                        }
                        trans.Commit();
                    }
                }
                info += "\n\t" + prompt;
                TaskDialog.Show("T: Revit", info);
                

                foreach (ElementId elementid in beams)
                {
                    
                }
            }
            catch (Exception e)
            {
                messages = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        public string GetParameterValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    return parameter.AsValueString();
                case StorageType.ElementId:
                    return parameter.AsElementId().IntegerValue.ToString();
                case StorageType.Integer:
                    return parameter.AsValueString();
                case StorageType.None:
                    return parameter.AsValueString();
                case StorageType.String:
                    return parameter.AsString();
                default:
                    return "";
            }
        }
    }
}
