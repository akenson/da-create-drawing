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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;

using Newtonsoft.Json;

using File = System.IO.File;
using Path = System.IO.Path;
using Directory = System.IO.Directory;
using DirectoryInfo = System.IO.DirectoryInfo;
using Newtonsoft.Json.Linq;

namespace CreateDrawingPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public void Run(Document placeholder)
        {
            LogTrace("Creating Drawing from iLogic rule...");
            string currDir = Directory.GetCurrentDirectory();

            LogTrace("currDir: " + currDir);
            // For local debugging
            //string inputPath = System.IO.Path.Combine(currDir, @"../../inputFiles", "params.json");
            //Dictionary<string, string> options = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText(inputPath));

            Dictionary<string, string> options = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText(Path.Combine(currDir,"inputParams.json")));
            string inputFile = options["inputFile"];
            string projectFile = options["projectFile"];
            string rule = options["runRule"];
            string drawingDocName = "result";

            string assemblyPath = Path.GetFullPath(Path.Combine(currDir, inputFile));

            string fullProjectPath = Path.GetFullPath(Path.Combine(currDir, projectFile));


            Console.WriteLine("fullProjectPath = " + fullProjectPath);

            DesignProject dp = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
            dp.Activate();

            Console.WriteLine("assemblyPath = " + assemblyPath);
            Document doc = inventorApplication.Documents.Open(assemblyPath);

            //RunRule(doc, rule);
            CreateDrawing(doc);

            // Drawing will be the last one created
            int docCount = inventorApplication.Documents.Count;
            Document lastDoc = inventorApplication.Documents[docCount];
            LogTrace("lastDoc: " + lastDoc.DisplayName);
            
            string drawingPath = Directory.GetCurrentDirectory() + "/result.idw";
            LogTrace("saving drawing: " + drawingPath);
            lastDoc.SaveAs(drawingPath, false);
            SaveAsPdf(lastDoc, drawingDocName);
        }

        private void CreateDrawing(Document doc)
        {
            double viewScale = 0.05;
            string templateLoc = "/Autodesk/Skid Packaging Layout.idw";

            // This gets the working directory of the Assembly
            DirectoryInfo parentDir = Directory.GetParent(Path.GetFullPath(doc.FullFileName));

            // Need one more directory up to get to the templates
            DirectoryInfo baseDir = Directory.GetParent(parentDir.FullName);

            string templateFile = baseDir.FullName + templateLoc;

            LogTrace("Adding Drawing template: " + templateFile);
            DrawingDocument drawingDoc = (DrawingDocument) inventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, templateFile);
            Sheet sheet = drawingDoc.Sheets[1];

            Point2d point1 = inventorApplication.TransientGeometry.CreatePoint2d(80, 40); // front view
            Point2d point2 = inventorApplication.TransientGeometry.CreatePoint2d(21, 29); // top view

            LogTrace("Adding Drawing Views...");
            DrawingView baseView = sheet.DrawingViews.AddBaseView((_Document)doc, point1, viewScale, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle, "My View");
            DrawingView projectedView = sheet.DrawingViews.AddProjectedView(baseView, point2, DrawingViewStyleEnum.kShadedDrawingViewStyle, viewScale);
        }


        private void SaveAsPdf(Document doc, string fileName)
        {
            string dirPath = Directory.GetCurrentDirectory();
            TranslatorAddIn oPDF = null;

            foreach (ApplicationAddIn item in inventorApplication.ApplicationAddIns)
            {

                if (item.ClassIdString == "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
                {
                    Trace.TraceInformation("Found the PDF addin.");
                    oPDF = (TranslatorAddIn)item;
                    break;
                }
                else { }
            }

            if (oPDF != null)
            {
                TranslationContext oContext = inventorApplication.TransientObjects.CreateTranslationContext();
                NameValueMap oPdfMap = inventorApplication.TransientObjects.CreateNameValueMap();

                if (oPDF.get_HasSaveCopyAsOptions(doc, oContext, oPdfMap))
                {
                    Trace.TraceInformation("PDF: can be exported.");

                    Trace.TraceInformation("PDF: Set context type");
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                    Trace.TraceInformation("PDF: create data medium");
                    DataMedium oData = inventorApplication.TransientObjects.CreateDataMedium();

                    string pdfFileName = dirPath + "/" + fileName + ".pdf";
                    Trace.TraceInformation("PDF save to: " + pdfFileName);
                    oData.FileName = pdfFileName;

                    oPdfMap.set_Value("All_Color_AS_Black", 0);

                    oPDF.SaveCopyAs(doc, oContext, oPdfMap, oData);
                    Trace.TraceInformation("PDF exported.");
                }

            }
        }

        public void RunWithArguments(Document doc, NameValueMap map)
        {
            LogTrace("RunWithArguments not implemented");
        }

        public void RunRule(Document doc, string rule)
        {
            string iLogicAddinGuid = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
            Inventor.ApplicationAddIn addin = null;
            try
            {
                addin = inventorApplication.ApplicationAddIns.get_ItemById(iLogicAddinGuid);
            }
            catch (Exception e)
            {
                LogError("Unable to load iLogic add-in: " + e.Message);
            }

            try
            {
                if (addin != null)
                {
                    LogTrace("Running rule: " + rule);
                    var iLogicAutomation = addin.Automation;
                    iLogicAutomation.RunRule(doc, rule);
                }
            }
            catch (Exception e)
            {
                LogError("Unable to run rule \"" + rule + "\": " + e.Message);
            }
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

        #endregion
    }
}