//
//////////////////////////////////////////////////////////////////////////////
//
//  Copyright 2015 Autodesk, Inc.  All rights reserved.
//
//  Use of this software is subject to the terms of the Autodesk license 
//  agreement provided at the time of installation or download, or which 
//  otherwise accompanies this software in either electronic or hard copy form.   
//
//////////////////////////////////////////////////////////////////////////////
// if just one type of hose exists, shortdescription should be "HOSE"


using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;


using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.AutoCAD.EditorInput;

using System;
using System.Runtime.InteropServices;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

using System.Collections.Generic;
using System.Reflection;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.P3dUI;
using Autodesk.ProcessPower.PartsRepository;
using System.IO;

namespace hose
{
    /// <summary>
    /// Helper class including some static helper functions.
    /// </summary>
    /// 


    public class Helper
    {
        public static Project currentProject { get; set; }
        public static Document ActiveDocument { get; set; }
        public static DataLinksManager ActiveDataLinksManager { get; set; }
        public static Database oDatabase { get; set; }
        public static Editor oEditor { get; set; }

        public static bool Initialize()
        {
            if (PlantApplication.CurrentProject == null)
                return false;

            currentProject = PlantApp.CurrentProject.ProjectParts["Piping"];
            ActiveDataLinksManager = currentProject.DataLinksManager;
            ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            oDatabase = ActiveDocument.Database;
            oEditor = ActiveDocument.Editor;
            return true;
        }

        public static void Terminate()
        {
            currentProject = null;
            ActiveDataLinksManager = null;
            ActiveDocument = null;
            oDatabase = null;
            oEditor = null;
        }

        [DllImport("accore.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedCmdS")]
        public static extern int acedCmdS(System.IntPtr vlist);

       
        public static PnPRow getPartSpecRow(PipePartSpecification pps, NominalDiameter nd, String shortdesc)
        {
            if (pps == null)
                return null;
            PnPTable table = pps.Database.Tables["EngineeringItems"];
            if (table != null)
            {
                String query = "\"NominalDiameter\"=" + nd.Value;
                //query += " and \"NominalUnit\"='" + part.PartSizeProperties.NominalDiameter.Units + "'";
                //query += " and \"PartCategory\"='" + part.PartSizeProperties.PropValue("PartCategory") + "'";
                if (shortdesc.Equals("")) shortdesc = "HOSE";
                query += " and \"ShortDescription\"='" + shortdesc + "'";



                try
                {
                    PnPRow[] r = table.Select(query);
                    oEditor.WriteMessage("2");
                    Helper.oEditor.WriteMessage(r.Length.ToString());
                    if (r.Length > 0)
                    {

                        return r[0];
                    }

                }
                catch (Exception e)
                {
                    oEditor.WriteMessage(e.Message);
                }



                /*catch (Autodesk.ProcessPower.DataObjects.Expression.PnPExpressionException e)
                {
                    oEditor.WriteMessage(e.Message);
                }
                catch (System.Runtime.InteropServices.SEHException e)
                {
                    oEditor.WriteMessage(e.Message);
                }*/
            }
            return null;
        }

        public static PipePartSpecification getPipePartSpec(String alternatepath)
        {

            UISettings sett = new UISettings();
            String specName = sett.CurrentSpec;
            PipePartSpecification cachePPS = null;
            PlantProject currentProj = PlantApplication.CurrentProject;
            String pathSpec = currentProj.ProjectFolderPath + "\\Spec Sheets\\" + specName + ".pspx";

            if (!File.Exists(pathSpec)) pathSpec = alternatepath;

            oEditor.WriteMessage("Using spec file: " + pathSpec + "\n");

            try
            {
                PipePartSpecification pps = PipePartSpecification.OpenSpecification(pathSpec);
                return cachePPS = pps;
            }
            catch (System.Exception e)
            {
                oEditor.WriteMessage("Error on open " + pathSpec);
                return null;
            }

        }


    }

    // Helper class to workaround a Hashtable issue: 
    // Can't change values in a foreach loop or enumerator
    class CBoolClass
    {
        public CBoolClass(bool val) { this.val = val; }
        public bool val;
        public override string ToString() { return (val.ToString()); }
    }
}

