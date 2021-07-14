using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.P3dUI;
using Autodesk.ProcessPower.PartsRepository;
using Autodesk.ProcessPower.PnP3dPlaceholderUtil;
using Autodesk.ProcessPower.PnP3dObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections;
using System.Collections.Specialized;
using Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(hose.mb_hose))]
[assembly: ExtensionApplication(typeof(hose.mb_hose))]

//version leer
//version 0.1, pfad war nicht allgemeingueltig, tool funktionierte nicht
//0.2 length rounded

namespace hose
{

    class mb_hose : IExtensionApplication
    {

        public void Initialize()
        {
            Helper.Initialize();
            Helper.oEditor.WriteMessage("hose0.2-usage:command hose \"shortdesc=Myhose,shape=1,scale=1,specalternatepath=C:/not/default/place/for/spec sheets\"");
            Helper.Terminate();
        }
        public void Terminate()
        {

        }



        [CommandMethod("hose", CommandFlags.UsePickSet)]
        public static void Do_mb_hose()
        {
            try
            {
                Helper.Initialize();

                /*if (Double.Parse(Helper.currentProject.Version) < 8)
                {
                    Helper.oEditor.WriteMessage("Versions earlier 2015 not supported! exit");
                    return;
                }*/

                object snapmodesaved = Application.GetSystemVariable("SNAPMODE");
                object osmodesaved = Application.GetSystemVariable("OSMODE");
                object os3dmodesaved = Application.GetSystemVariable("3DOSMODE");

                Application.SetSystemVariable("SNAPMODE", 0);
                Application.SetSystemVariable("OSMODE", 0);
                Application.SetSystemVariable("3DOSMODE", 0);

                string configstr = "";
                PromptResult pr = Helper.oEditor.GetString("\nconfiguration string: ");
                if (pr.Status != PromptStatus.OK)
                {
                    Helper.oEditor.WriteMessage("No configuration string was provided, using defaults\n");
                }
                else
                    configstr = pr.StringResult;

                string shortdesc = "Hose";
                string shape = "0";
                string scalefactor = "1";
                double thescalefactor = 1.0;
                string alternatepath = "";
                int linearUnit = Autodesk.ProcessPower.AcPp3dObjectsUtils.ProjectUnits.CurrentLinearUnit; //Unit drawing
                int ndUnit = Autodesk.ProcessPower.AcPp3dObjectsUtils.ProjectUnits.CurrentNDUnit; //Unit nominal diameter display
                string theunit = "";

                if (linearUnit == 1)
                    theunit = "mm";
                else
                    theunit = "inch";

                /* Units: 1 = Metric
                          2 = Imperial
             
             
                 command hose "shortdesc=Hose123,shape=3,scale=1,specalternatepath=C:/Users/bussm/AppData/Roaming/Autodesk/Autodesk AutoCAD Plant 3D 2016/R20.1/deu/SampleProjects/SampleProject"
                 */

                if (!configstr.Equals(""))
                {
                    string[] configArr = configstr.Split(new char[] { ',' });
                    string tmpstr = "";
                    tmpstr = configArr[0].Split(new char[] { '=' })[1];
                    if (!tmpstr.Trim().Equals("")) shortdesc = tmpstr;
                    tmpstr = configArr[1].Split(new char[] { '=' })[1];
                    if (!tmpstr.Trim().Equals("")) shape = tmpstr;
                    tmpstr = configArr[2].Split(new char[] { '=' })[1];
                    if (!tmpstr.Trim().Equals("")) scalefactor = tmpstr;
                    tmpstr = configArr[3].Split(new char[] { '=' })[1];
                    if (!tmpstr.Trim().Equals("")) alternatepath = tmpstr;
                }

                try { thescalefactor = Double.Parse(scalefactor); }
                catch { }

                PromptEntityOptions opt = new PromptEntityOptions("Select the spline");
                PromptEntityResult res;
                Spline myspline = new Spline();
                Point3d point1 = new Point3d();
                Point3d point2 = new Point3d();
                Point3d point3 = new Point3d();
                Point3d point4 = new Point3d();

                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                double NDdouble = nd.Value;
                if (linearUnit == 1 && ndUnit == 2)
                    NDdouble = NDdouble * 25.4;

                //Helper.oEditor.WriteMessage("version: " + Helper.currentProject.Version);
                //9=2016
                //8=2015

                do
                {
                    res = Helper.oEditor.GetEntity(opt);
                    opt.Message = "Select the spline";
                }
                while (res.Status == PromptStatus.Error);
                if (res.Status == PromptStatus.OK)
                {
                    using (Transaction t = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
                    {
                        Entity ent = (Entity)t.GetObject(res.ObjectId, OpenMode.ForWrite);
                        myspline = (Spline)ent;


                        double thelen = myspline.GetDistanceAtParameter(myspline.EndParam) - myspline.GetDistanceAtParameter(myspline.StartParam);
                        //Helper.oEditor.WriteMessage("/nthelen: " + thelen);
                        int numfits = myspline.NumFitPoints;
                        for (int i = 0; i < numfits; i++)
                        {
                            if (i == 0)
                                point1 = myspline.GetFitPointAt(i);
                            else if (i == 1)
                                point2 = myspline.GetFitPointAt(i);
                            else if (i == numfits - 2)
                                point3 = myspline.GetFitPointAt(numfits - 2);
                            else if (i == numfits - 1)
                                point4 = myspline.GetFitPointAt(numfits - 1);
                        }

                        Vector3d theVector = point1.GetVectorTo(point2);
                        Vector3d theMainVector = theVector.GetPerpendicularVector();
                        theMainVector = theMainVector.MultiplyBy(NDdouble * thescalefactor / 2);

                        Curve myshape;

                        if (shape.Equals("0") || shape.Equals("1") || shape.Equals("2"))
                        {
                            Ellipse myellipse = new Ellipse(point1, theVector, theMainVector, 0.8, 0.0, 2 * Math.PI);
                            myellipse.SetDatabaseDefaults(Helper.oDatabase);
                            myshape = myellipse;
                        }
                        else
                        {
                            Circle mycircle = new Circle(point1, theVector, NDdouble * thescalefactor / 2);
                            mycircle.SetDatabaseDefaults(Helper.oDatabase);
                            myshape = mycircle;
                        }



                        BlockTableRecord btr = (BlockTableRecord)t.GetObject(Helper.oDatabase.CurrentSpaceId, OpenMode.ForWrite);

                        Solid3d mysolid = new Solid3d();

                        SweepOptionsBuilder theSweepOptions = new SweepOptionsBuilder();

                        if (shape.Equals("0"))
                            theSweepOptions.TwistAngle = thelen / NDdouble;
                        else if (shape.Equals("1"))
                            theSweepOptions.TwistAngle = 2 * thelen / NDdouble;
                        else if (shape.Equals("2"))
                            theSweepOptions.TwistAngle = 4 * thelen / NDdouble;

                        mysolid.CreateSweptSolid(myshape, myspline, theSweepOptions.ToSweepOptions());

                        ObjectId solidID = btr.AppendEntity(mysolid);

                        t.AddNewlyCreatedDBObject(mysolid, true);

                        myspline.Erase();



                        //Helper.oEditor.WriteMessage(myspline.NumFitPoints.ToString());

                        //replace blockname with timeinmillis

                        long thestamp = System.Diagnostics.Stopwatch.GetTimestamp();


                        Helper.oEditor.Command("_-block", thestamp.ToString(), point1, solidID, "");
                        Helper.oEditor.Command("_-insert", '"' + thestamp.ToString() + '"', point1, 1, 1, 0);
                        PromptSelectionResult whatsLast = Helper.oEditor.SelectLast();
                        ObjectId lastObject = whatsLast.Value[0].ObjectId;
                        Helper.oEditor.Command("PLANTPARTCONVERT", "_l", "_a", point1, point2, "_f", "_a", "_a", point4, point3, "_f", "_a", "_x");

                        Entity cblock = (Entity)t.GetObject(lastObject, OpenMode.ForWrite);
                        cblock.Erase();

                        PipePartSpecification thespec = Helper.getPipePartSpec(alternatepath);
                        PnPRow specrow = Helper.getPartSpecRow(thespec, nd, shortdesc);

                        ArrayList portlist = new ArrayList(2);
                        Dictionary<string, string> portS1 = new Dictionary<string, string>();
                        Dictionary<string, string> portS2 = new Dictionary<string, string>();
                        Dictionary<string, string> propdict = new Dictionary<string, string>();

                        if (specrow != null)
                        {

                            portS1.Add("PortName", "S1");
                            portS1.Add("NominalDiameter", nd.Value.ToString());
                            portS1.Add("NominalUnit", specrow["NominalUnit"].ToString());
                            portS1.Add("MatchingPipeOd", specrow["MatchingPipeOd"].ToString());
                            portS1.Add("EndType", specrow["EndType"].ToString());
                            portS1.Add("FlangeStd", specrow["FlangeStd"].ToString());
                            portS1.Add("GasketStd", specrow["GasketStd"].ToString());
                            portS1.Add("Facing", specrow["Facing"].ToString());
                            portS1.Add("FlangeThickness", specrow["FlangeThickness"].ToString());
                            portS1.Add("PressureClass", specrow["PressureClass"].ToString());
                            portS1.Add("Schedule", specrow["Schedule"].ToString());
                            portS1.Add("WallThickness", specrow["WallThickness"].ToString());
                            portS1.Add("EngagementLength", specrow["EngagementLength"].ToString());
                            portS1.Add("LengthUnit", specrow["LengthUnit"].ToString());


                            portS2.Add("PortName", "S2");
                            portS2.Add("NominalDiameter", nd.Value.ToString());
                            portS2.Add("NominalUnit", specrow["NominalUnit"].ToString());
                            portS2.Add("MatchingPipeOd", specrow["MatchingPipeOd"].ToString());
                            portS2.Add("EndType", specrow["EndType"].ToString());
                            portS2.Add("FlangeStd", specrow["FlangeStd"].ToString());
                            portS2.Add("GasketStd", specrow["GasketStd"].ToString());
                            portS2.Add("Facing", specrow["Facing"].ToString());
                            portS2.Add("FlangeThickness", specrow["FlangeThickness"].ToString());
                            portS2.Add("PressureClass", specrow["PressureClass"].ToString());
                            portS2.Add("Schedule", specrow["Schedule"].ToString());
                            portS2.Add("WallThickness", specrow["WallThickness"].ToString());
                            portS2.Add("EngagementLength", specrow["EngagementLength"].ToString());
                            portS2.Add("LengthUnit", specrow["LengthUnit"].ToString());


                            propdict.Add("ShortDescription", specrow["ShortDescription"].ToString());
                            propdict.Add("PartSizeLongDesc", specrow["PartSizeLongDesc"].ToString() + " " + Math.Round(thelen) + "" + theunit);
                            propdict.Add("PartFamilyLongDesc", specrow["PartFamilyLongDesc"].ToString());
                            propdict.Add("CompatibleStandard", specrow["CompatibleStandard"].ToString());
                            propdict.Add("Manufacturer", specrow["Manufacturer"].ToString());
                            propdict.Add("ItemCode", specrow["ItemCode"].ToString());
                            propdict.Add("DesignStd", specrow["DesignStd"].ToString());
                            propdict.Add("DesignPressureFactor", specrow["DesignPressureFactor"].ToString());
                            propdict.Add("Weight", specrow["Weight"].ToString());
                            propdict.Add("WeightUnit", specrow["WeightUnit"].ToString());
                            propdict.Add("ContentIsoSymbolDefinition", specrow["ContentIsoSymbolDefinition"].ToString());
                            propdict.Add("PartCategory", specrow["PartCategory"].ToString());
                            propdict.Add("ComponentDesignation", specrow["ComponentDesignation"].ToString());
                            propdict.Add("PartVersion", specrow["PartVersion"].ToString());
                            propdict.Add("Status", specrow["Status"].ToString());

                        }
                        else
                        {
                            Helper.oEditor.WriteMessage("no part information in the spec, using default data");
                            portS1.Add("PortName", "S1");
                            portS1.Add("NominalDiameter", nd.Value.ToString());
                            portS1.Add("NominalUnit", theunit);
                            portS1.Add("EndType", "PL");
                            portS2.Add("PortName", "S2");
                            portS2.Add("NominalDiameter", nd.Value.ToString());
                            portS2.Add("NominalUnit", theunit);
                            portS2.Add("EndType", "PL");
                        }

                        portlist.Add(portS1);
                        portlist.Add(portS2);

                        /*foreach (KeyValuePair<string, string> entry in propdict)
                        {
                            Helper.oEditor.WriteMessage(entry.Key + "-" + entry.Value + "\n");
                        }*/

                        t.Commit();

                        Application.SetSystemVariable("SNAPMODE", snapmodesaved);
                        Application.SetSystemVariable("OSMODE", osmodesaved);
                        Application.SetSystemVariable("3DOSMODE", os3dmodesaved);

                        //InsertCustomGeometryPlaceHolderInModel(string dwgPath, string blockName, string partType, string partFamilyId, Autodesk.ProcessPower.PartsRepository.NominalDiameter nd, string Tag, bool TagAtInsert, bool SpecAtInsert, bool bPlaceHolder, System.Collections.Generic.Dictionary<string, string> Properties, System.Collections.ArrayList portProperties);    
                        
                        PnP3dPlaceholderUtil.InsertCustomGeometryPlaceHolderInModel(Helper.ActiveDocument.Name, ("" + thestamp), "Coupling", "", new NominalDiameter(100), "", false, false, true, propdict, portlist);


                    }
                }



            }
            catch (System.Exception e)
            {

                Helper.oEditor.WriteMessage("/nerror: " + e.Message.ToString());

            }
            finally
            {

                Helper.Terminate();

            }
        }
    }
}
