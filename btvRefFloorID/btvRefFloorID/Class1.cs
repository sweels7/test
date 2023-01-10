using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Document = Autodesk.Revit.DB.Document;

namespace btvRefFloorID
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        UIApplication app;
        Document doc;

        public Result Execute(ExternalCommandData commandData, ref string messages, ElementSet elements)
        {
            app = commandData.Application;
            doc = app.ActiveUIDocument.Document;

            /************************在專案新增共用參數BTV_Ref_Floor_ID************************/
            StringBuilder levelOrderInformation2 = new StringBuilder();
            Category category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFraming);
            CategorySet categorySet = app.Application.Create.NewCategorySet();
            categorySet.Insert(category);
            string originalFile = app.Application.SharedParametersFilename;
            string tempFile = @"C:\btvRefFloorID.txt";
            //string tempFile = "Data source=myData.db";

            app.Application.SharedParametersFilename = tempFile;
            DefinitionFile shareParameterFile = app.Application.OpenSharedParameterFile();
            try
            {
                foreach (DefinitionGroup dg in shareParameterFile.Groups)
                {
                    if (dg.Name == "GROUP1")
                    {
                        ExternalDefinition externalDefinition = dg.Definitions.get_Item("BTV_Ref_Floor_ID") as ExternalDefinition;
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
            catch(Exception e)
            {
                messages = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
            finally
            {
                app.Application.SharedParametersFilename = originalFile;
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}
