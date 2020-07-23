/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.IO.Compression;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ShrinkWrapPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public Dictionary<string, dynamic> CreateEnumMap()
        {
            Type[] enumTypes = new Type[]
            {
                typeof(DerivedComponentStyleEnum),
                typeof(ShrinkwrapRemoveStyleEnum)
            };

            var map = new Dictionary<string, dynamic>();
            foreach (var enumType in enumTypes)
            {
                foreach (var value in Enum.GetValues(enumType))
                {
                    var name = Enum.GetName(enumType, value);
                    map.Add(name, value);
                }
            }

            return map;
        }

        private T GetValue<T>(JObject obj, string name)
        {
            JObject val = obj.GetValue(name) as JObject;
            return val.GetValue("value").Value<T>();
        }

        //public static class TranslatorAddins
        //{
        //    public const string SAT	= "89162634-02B6-11D5-8E80-0010B541CD80";
        //    public const string STEP = "90AF7F40-0C01-11D5-8E83-0010B541CD80";
        //    public const string IGES = "90AF7F44-0C01-11D5-8E83-0010B541CD80";
        //    public const string CATIAV5ProductExport = "8A88FC01-0C32-4B3E-BE12-DDC8DF6FFF18";
        //    public const string DWG	= "C24E3AC2-122E-11D5-8E91-0010B541CD80";
        //    public const string DXF	= "C24E3AC4-122E-11D5-8E91-0010B541CD80";
        //    public const string STLImport = "81CA7D27-2DBE-4058-8188-9136F85FC859";
        //    public const string DWF	= "0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4";
        //    public const string PDF	= "0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4";
        //    public const string DWFx	= "0AC6FD97-2F4D-42CE-8BE0-8AEA580399E4";
        //    public const string CATIAV5PartExport = "2FEE4AE5-36D3-4392-89C7-58A9CD14D305";
        //    public const string RVT	= "2058EF4F-37A3-4B57-A322-B4E79E7D53E4";
        //    public const string TUFF = "3260333F-3B0D-4812-9274-E94E14A77A16";
        //    public const string ParasolidText = "8F9D3571-3CB8-42F7-8AFF-2DB2779C8465";
        //    public const string Fusion	= "C6B37B88-3CFA-4521-9873-E087B8626C44";
        //    public const string FCADTransServer = "BE52A5E7-58D8-4E3C-A887-06A4C8F29568";
        //    public const string AutodeskIDFTranslator = "6C5BBC04-5D6F-4353-94B1-060CD6554444";
        //    public const string SolidWorks = "402BE503-725D-41CB-B746-D557AB83BAF1";
        //    public const string ProENGINEERGranite = "66CB2667-73AD-401C-A531-64EC701825A1";
        //    public const string NX = "93D506C4-8355-4E28-9C4E-C2B5F1EDC6AE";
        //    public const string SMT	= "B4ECC5EB-9507-46E5-87FB-EBB9479CE1DF";
        //    public const string OBJImport = " C420F7E4-98FD-4A57-BC1E-04D1D683EFDF";
        //    public const string SVF	= "C200B99B-B7DD-4114-A5E9-6557AB5ED8EC";
        //    public const string ParasolidBinary = " A8F8F8E5-BBAB-4F74-8B1B-AC011251F8AC";
        //    public const string ProENGINEERandCreoParametric = "46D96B7A-CF8A-49C9-8703-2F40CFBDF547";
        //    public const string ContentCenterItemTranslator = "A547F528-D239-475F-8FC6-8F97C4DB6746";
        //    public const string SolidEdge = " E2548DAF-D56B-4809-82B9-5F670E6D518B";
        //    public const string ProENGINEERNeutral = "8CEC09E3-D638-4E8F-A6E1-0D1E1A5FC8E3";
        //    public const string CATIAV4Import	= "C6ACD948-E1C5-4B5B-ADEE-3ED968F8CB1A";
        //    public const string Rhino = "2CB23BF0-E2AC-4B32-B0A1-1CC292AF6623";
        //    public const string CATIAV5Import	= "8D1717FA-EB24-473C-8B0F-0F810C4FC5A8";
        //    public const string AutodeskDWFMarkupManager = "55EBD0FA-EF60-4028-A350-502CA148B499";
        //    public const string JT	= "16625A0E-F58C-4488-A969-E7EC4F99CACD";
        //    public const string Alias	= "DC5CD10A-F6D1-4CA3-A6E3-42A6D646B03E";
        //    public const string OBJExport = " F539FB09-FC01-4260-A429-1818B14D6BAC";
        //    public const string STLExport = "533E9A98-FC3B-11D4-8E7E-0010B541CD80";
        //}

        //public AssemblyDocument Import(string fullAssemblyPath)
        //{
        //    LogTrace("Starting Import");
        //    ApplicationAddIns addins = inventorApplication.ApplicationAddIns;
        //    //foreach (ApplicationAddIn addin in addins)
        //    //{
        //    //    if  (addin is )
        //    //}
        //    //If TypeOf addin Is TranslatorAddIn Then
        //    //    LogTrace(addin.DisplayName & Chr(9) & addin.ClassIdString)
        //    //End If
        //    //Next
        //    TranslatorAddIn transAddin = (TranslatorAddIn)addins.ItemById["{" + TranslatorAddins.STEP + "}"];
        //    LogTrace("Get Stp Addin");
        //    TransientObjects transObjects = inventorApplication.TransientObjects;
        //    transAddin.Activate();

        //    LogTrace("Stp Addin Activated");

        //    //Prepare the 1st parameter for Open(), the file name
        //    DataMedium file = transObjects.CreateDataMedium();
        //    file.FileName = fullAssemblyPath;

        //    LogTrace("DataMedium created");

        //    //Prepare the 2nd parameter for Open(), the open type, for convenience set it as drag&drop
        //    TranslationContext context = transObjects.CreateTranslationContext();
        //    context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

        //    LogTrace("TranslationContext created");

        //    //Prepare the 3rd parameter for Open(), the import options
        //    NameValueMap options = transObjects.CreateNameValueMap();

        //    Boolean hasOpt = transAddin.HasOpenOptions[file, context, options];
        //    options.Value["AssociativeImport"] = false;
        //    options.Value["EmbedInDocument "] = false;
        //    options.Value["SaveLocationIndex "] = 0;
        //    options.Value["TessellationDetailIndex"] = 1;

        //    for (int i = 1; i <= options.Count; ++i)
        //    {
        //        string name = options.Name[i];
        //        string value = (string)options.Value[name];
        //        LogTrace(name + " " + value);
        //    }

        //    object sourceObject;
        //    transAddin.Open(file, context, options, out sourceObject);

        //    AssemblyDocument doc = (AssemblyDocument)sourceObject;

        //    return doc;
        //}


        public AssemblyDocument AnyCADImport(string fullAssemblyPath)
        {
            var em = inventorApplication.ErrorManager;
            LogTrace("AnyCADImport called");
            inventorApplication.SaveOptions.TranslatorReportLocation = ReportLocationEnum.kNoReport;
            AssemblyDocument doc = (AssemblyDocument)inventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
            LogTrace("New Assembly created");
            AssemblyComponentDefinition compDef = doc.ComponentDefinition;
            ImportedGenericComponentDefinition importedCompDef = (ImportedGenericComponentDefinition)compDef.ImportedComponents.CreateDefinition(fullAssemblyPath);
            importedCompDef.ReferenceModel = true;

            string saveFilesLocation = System.IO.Path.GetDirectoryName(fullAssemblyPath) + "\\";
            LogTrace("Imported Files Location = " + saveFilesLocation);
            importedCompDef.SaveFilesLocation = saveFilesLocation;
            ImportedComponent importedComp = compDef.ImportedComponents.Add((ImportedComponentDefinition)importedCompDef);

            try
            {
                LogTrace("Before Update");
                doc.Update2();
                LogTrace("After Update");
                LogTrace("Before Save");

                LogTrace($"Checking ErrorManager before save");
                LogTrace($"HasErrors = {em.HasErrors}");
                LogTrace($"HasWarnings = {em.HasWarnings}");
                LogTrace($"AllMessages = {em.AllMessages}");

                doc.SaveAs(System.IO.Path.Combine(saveFilesLocation, "output.iam"), false);
                LogTrace("After Save");

                LogTrace("Assembly Path = " + doc.FullFileName);

                LogTrace($"Checking ErrorManager after save");
                LogTrace($"HasErrors = {em.HasErrors}");
                LogTrace($"HasWarnings = {em.HasWarnings}");
                LogTrace($"AllMessages = {em.AllMessages}");
            }
            catch (Exception ex)
            {
                LogTrace("Error: " + ex.Message);
            }
            LogTrace("Successfully Imported");
            return doc;
        }

        //public void SetExternalRulesDirectory(string externalRuleDirName)
        //{
        //    LogTrace(" ************ Start Processing Adding External Rules Directory *********");
        //    string ClientId = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
        //    ApplicationAddIn iLogicAddIn = inventorApplication.ApplicationAddIns.ItemById[(ClientId)];
        //    dynamic iLogicAuto = iLogicAddIn.Automation;
        //    if (iLogicAuto == null)
        //    {
        //        LogError("Error: No iLogicAuto Atuomation set!");
        //    }
        //    else
        //    {
        //        List<string> externalRuleDirectories = new List<string>();
        //        foreach (string path in iLogicAuto.FileOptions.ExternalRuleDirectories)
        //        {
        //            externalRuleDirectories.Add(path);
        //            LogTrace("Existing Extrenal Rule Directory:" + path);
        //        }
        //        if (!externalRuleDirectories.Contains(externalRuleDirName))
        //        {
        //            LogTrace("Add a new Extrenal Directory:" + externalRuleDirName);
        //            externalRuleDirectories.Add(externalRuleDirName);
        //            iLogicAuto.FileOptions.ExternalRuleDirectories = externalRuleDirectories.ToArray();
        //        }
        //    }
        //    LogTrace(" ************ End Processing Adding External Rules Directory *********");
        //}

        public void Run(Document doc)
        {
            LogTrace("Run()");

            string currentDir = System.IO.Directory.GetCurrentDirectory();
            LogTrace("Current Dir = " + currentDir);

            var map = CreateEnumMap();

            JObject parameters = JObject.Parse(System.IO.File.ReadAllText("params.json"));

            string assemblyPath = parameters.GetValue("assemblyPath").Value<string>();

            using (new HeartBeat())
            {
                
                CheckPerformance();
                if (parameters.ContainsKey("projectPath"))
                {
                    string projectPath = parameters.GetValue("projectPath").Value<string>();
                    string fullProjectPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, projectPath));
                    LogTrace("fullProjectPath = " + fullProjectPath);
                    
                    if (System.IO.File.Exists(fullProjectPath))
                    {
                        LogTrace("Loading and activating project");
                        DesignProject dp = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
                        dp.Activate();
                    }
                    else
                    {
                        LogTrace("Project file not found");
                    }
                }

                string fullAssemblyPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, assemblyPath));
                LogTrace("fullAssemblyPath = " + fullAssemblyPath);
                if (!System.IO.File.Exists(fullAssemblyPath))
                {
                    LogTrace("Did not find assembly");
                    return;
                }

                string ext = System.IO.Path.GetExtension(fullAssemblyPath);
                AssemblyDocument asmDoc;
                if (ext == ".iam" )
                {
                    if (parameters.ContainsKey("LOD"))
                    {
                        fullAssemblyPath += "<" + parameters.GetValue("LOD").Value<string>() + ">";
                    }
                    asmDoc = inventorApplication.Documents.Open(fullAssemblyPath, true) as AssemblyDocument;
                }
                else
                {
                    asmDoc = AnyCADImport(fullAssemblyPath);
                    fullAssemblyPath = System.IO.Path.ChangeExtension(fullAssemblyPath, ".iam");
                }

                LogTrace("fullAssemblyPath = " + fullAssemblyPath);
                LogTrace("Opened input assembly file");
                AssemblyComponentDefinition compDef = asmDoc.ComponentDefinition;

                LogTrace("Before Update2");
                asmDoc.Update2(true);
                LogTrace("After Update2");

                PartDocument partDoc = inventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", true) as PartDocument;
                LogTrace("Created part document");
                PartComponentDefinition partCompDef = partDoc.ComponentDefinition;

                ShrinkwrapDefinition SWD = null;
                try
                {
                    LogTrace("asmDoc.FullDocumentName = " + asmDoc.FullDocumentName);
                    LogTrace("LOD = " + asmDoc.LevelOfDetailName);
                    SWD = partCompDef.ReferenceComponents.ShrinkwrapComponents.CreateDefinition(asmDoc.FullDocumentName);
                    LogTrace("After ShrinkwrapComponents.CreateDefinition");
                    if (parameters.ContainsKey("CreateIndependentBodiesOnFailedBoolean"))
                    {
                        SWD.CreateIndependentBodiesOnFailedBoolean =
                            GetValue<bool>(parameters, "CreateIndependentBodiesOnFailedBoolean");
                    }
                    if (parameters.ContainsKey("DeriveStyle"))
                    {
                        // e.g. DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
                        SWD.DeriveStyle = map[GetValue<string>(parameters, "DeriveStyle")];
                    }
                    if (parameters.ContainsKey("RemoveInternalParts"))
                    {
                        SWD.RemoveInternalParts = GetValue<bool>(parameters, "RemoveInternalParts");
                    }
                    if (parameters.ContainsKey("RemoveAllInternalVoids"))
                    {
                        SWD.RemoveAllInternalVoids = GetValue<bool>(parameters, "RemoveAllInternalVoids");
                    }
                    if (parameters.ContainsKey("RemoveHolesDiameterRange"))
                    {
                        SWD.RemoveHolesDiameterRange = GetValue<double>(parameters, "RemoveHolesDiameterRange"); // in cm
                    }
                    if (parameters.ContainsKey("RemoveHolesStyle"))
                    {
                        SWD.RemoveHolesStyle = map[GetValue<string>(parameters, "RemoveHolesStyle")];
                    }
                    if (parameters.ContainsKey("RemoveFilletsRadiusRange"))
                    {
                        SWD.RemoveFilletsRadiusRange = GetValue<double>(parameters, "RemoveFilletsRadiusRange"); // in cm
                    }
                    if (parameters.ContainsKey("RemoveFilletsStyle"))
                    {
                        SWD.RemoveFilletsStyle = map[GetValue<string>(parameters, "RemoveFilletsStyle")];
                    }
                    if (parameters.ContainsKey("RemoveChamfersDistanceRange"))
                    {
                        SWD.RemoveChamfersDistanceRange = GetValue<double>(parameters, "RemoveChamfersDistanceRange");
                    }
                    if (parameters.ContainsKey("RemoveChamfersStyle"))
                    {
                        SWD.RemoveChamfersStyle = map[GetValue<string>(parameters, "RemoveChamfersStyle")];
                    }
                    if (parameters.ContainsKey("RemovePartsBySize"))
                    {
                        SWD.RemovePartsBySize = GetValue<bool>(parameters, "RemovePartsBySize");
                    }
                    if (parameters.ContainsKey("RemovePartsSize"))
                    {
                        SWD.RemovePartsSize = GetValue<double>(parameters, "RemovePartsSize"); // in cm
                    }
                    if (parameters.ContainsKey("RemovePocketsStyle"))
                    {
                        SWD.RemovePocketsStyle = map[GetValue<string>(parameters, "RemovePocketsStyle")];
                    }
                    if (parameters.ContainsKey("RemovePocketsMaxFaceLoopRange"))
                    {
                        SWD.RemovePocketsMaxFaceLoopRange = GetValue<double>(parameters, "RemovePocketsMaxFaceLoopRange"); // in cm
                    }

                    LogTrace("Before ShrinkwrapComponents.Add");
                    ShrinkwrapComponent SWComp = null;
                    try
                    {
                        SWComp = partCompDef.ReferenceComponents.ShrinkwrapComponents.Add(SWD);
                    }
                    catch (Exception ex)
                    {
                        LogTrace(ex.Message);
                        SWComp = partCompDef.ReferenceComponents.ShrinkwrapComponents[1];
                    }
                    LogTrace("After ShrinkwrapComponents.Add");

                    LogTrace("Before SuppressLinkToFile = true");
                    try
                    {
                        if (parameters.ContainsKey("SuppressLinkToFile"))
                        { 
                            SWComp.SuppressLinkToFile = parameters.GetValue("SuppressLinkToFile").Value<bool>();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogTrace(ex.Message);
                    }
                    LogTrace("After SuppressLinkToFile = true");

                    LogTrace("Saving part document");
                    partDoc.SaveAs(System.IO.Path.Combine(currentDir, "output.ipt"), false);
                    LogTrace("Saved part document to output.ipt");

                    //LogTrace("Saving to OBJ");
                    //partDoc.SaveAs(System.IO.Path.Combine(currentDir, "outputObjZip", "output.obj"), true);
                    //LogTrace("Saved to OBJ named output.obj");
                }
                catch (Exception ex)
                {
                    LogTrace("Error: " + ex.Message);
                }

                LogTrace("Finished");

            }
        }

        public void RunWithArguments(Document doc, NameValueMap map)
        {
            LogTrace("RunWithArguments()");

            Run(doc);
        }

        #region Logging utilities

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        private async Task CheckPerformance()
        {
            Console.WriteLine("Starting to Check Performance");
            while (true)
            {
                try
                {
                    var proc = Process.GetCurrentProcess();
                    var mem = proc.WorkingSet64;
                    var cpu = proc.TotalProcessorTime;
                    LogTrace("My process used working set {0:n3} K of working set and CPU {1:n} msec", mem / 1024.0, cpu.TotalMilliseconds);
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Failed Memory Check: " + exception.Message);
                }
                await Task.Delay(5000);
            }
        }

        #endregion
    }
}