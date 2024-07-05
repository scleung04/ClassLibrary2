using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;

namespace ClassLibrary2
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Path to the folder containing Revit files
            string folderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Revit Files";
            // Path to the log file
            string logFilePath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Dynamo_Revit_Log.txt";

            try
            {
                // Get the current Revit application
                UIApplication uiApp = commandData.Application;
                Application app = uiApp.Application;

                // Open the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    // Get all files in the folder
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    foreach (string filePath in files)
                    {
                        Document doc = null;

                        try
                        {
                            // Open the Revit file with failure handling options to suppress warnings
                            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                            OpenOptions openOptions = new OpenOptions();

                            // Set up the Failure Handling Options to ignore warnings
                            FailureHandlingOptions failureHandlingOptions = new FailureHandlingOptions();
                            failureHandlingOptions.SetClearAfterRollback(true);
                            failureHandlingOptions.SetFailuresPreprocessor(new WarningSuppressor());

                            // Register a failure processor before opening the document
                            FailureProcessor processor = new FailureProcessor(app, failureHandlingOptions);

                            doc = app.OpenDocumentFile(modelPath, openOptions);

                            // If the document was successfully opened
                            if (doc != null)
                            {
                                // Write success message to log file
                                writer.WriteLine(DateTime.Now.ToString() + ": Successfully opened Revit file: " + filePath);

                                // Check for the panel-circuit mismatch
                                bool mismatchFound = CheckPanelCircuitMismatch(doc, writer);

                                if (mismatchFound)
                                {
                                    writer.WriteLine(DateTime.Now.ToString() + ": Panel-circuit mismatch found. Skipping file: " + filePath);
                                    continue; // Skip the file if mismatch is found
                                }

                                // Export to IFC using Autodesk.IFC.Export.UI
                                ExportToIFC(doc, filePath, writer);
                            }
                            else
                            {
                                HandleFileOpeningError(filePath, writer);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Write detailed error message to log file
                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred with file " + filePath + ": " + ex.Message);
                            writer.WriteLine(ex.StackTrace);
                        }
                        finally
                        {
                            // Ensure the document is closed if it was opened
                            if (doc != null && doc.IsValidObject)
                            {
                                doc.Close(false); // Close without saving changes
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Log the operation cancellation
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + ": The operation was cancelled by the user.");
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + ": An error occurred: " + ex.Message);
                }
            }

            return Result.Succeeded;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer)
        {
            writer.WriteLine($"{DateTime.Now}: Failed to open Revit file: {filePath}");
        }

        private bool CheckPanelCircuitMismatch(Document doc, StreamWriter writer)
        {
            bool mismatchFound = false;

            var panels = new FilteredElementCollector(doc)
                         .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                         .OfClass(typeof(FamilyInstance))
                         .WhereElementIsNotElementType()
                         .OfType<FamilyInstance>()
                         .ToList();

            foreach (var panel in panels)
            {
                double totalLoad = panel.LookupParameter("Total Connected Load")?.AsDouble() ?? 0;
                double panelCapacity = panel.LookupParameter("Panel Capacity")?.AsDouble() ?? 0;

                if (totalLoad > panelCapacity)
                {
                    mismatchFound = true;
                    writer.WriteLine($"{DateTime.Now}: Mismatch found for panel {panel.Name}. Total load exceeds capacity.");
                }
            }

            return mismatchFound;
        }

        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                string ifcFolderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\R2024";
                Directory.CreateDirectory(ifcFolderPath);

                string ifcFileName = Path.GetFileNameWithoutExtension(filePath) + ".ifc";
                string ifcFilePath = Path.Combine(ifcFolderPath, ifcFileName);

                // Set up IFC export options
                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    FileVersion = IFCVersion.IFC2x3CV2,
                    ExportBaseQuantities = true,
                    SpaceBoundaryLevel = 2 // Level of space boundaries
                };

                // Set exchange requirements
                ifcOptions.AddOption("ExchangeRequirement", "IFC4_Reference_View"); // Example exchange requirement

                // Set file type
                ifcOptions.AddOption("FileType", "IFC2x3"); // Example file type setting

                // Set the phase to export
                PhaseArray phases = doc.Phases;
                Phase phaseToExport = phases.get_Item(phases.Size - 1); // Export the last phase
                ifcOptions.AddOption("ExportingPhase", phaseToExport.Id.IntegerValue.ToString());

                // Space boundaries
                ifcOptions.AddOption("SpaceBoundaries", "none"); // Example setting for space boundaries

                // Additional IFC export options
                ifcOptions.AddOption("SplitWallsAndColumnsByLevel", "false");

                // Tab 2: Additional Content
                ifcOptions.AddOption("ExportElementsVisibleInView", "false"); // Export only elements visible in the view
                ifcOptions.AddOption("ExportRooms", "false"); // Export rooms
                ifcOptions.AddOption("ExportAreas", "false"); // Export areas
                ifcOptions.AddOption("ExportSpaces", "false"); // Export spaces in 3D views
                ifcOptions.AddOption("ExportSteelElements", "true"); // Include steel elements
                ifcOptions.AddOption("Export2DPlanViewElements", "false"); // Export 2D plan view elements

                // Tab 3: Property Sets
                ifcOptions.AddOption("ExportRevitPropertySets", "false");
                ifcOptions.AddOption("ExportIFCCommonPropertySets", "true");
                ifcOptions.AddOption("ExportBaseQuantities", "false");
                ifcOptions.AddOption("ExportMaterialPropertySets", "false"); // Export material property sets
                ifcOptions.AddOption("ExportSchedulesAsPsets", "false");
                ifcOptions.AddOption("ExportOnlySchedulesContainingIFCPsetOrCommonInTitle", "false"); // Export only schedules containing IFC, Pset, or Common in the title
                ifcOptions.AddOption("ExportUserDefinedPsets", "false");
                ifcOptions.AddOption("ExportUserDefinedPsetsFile", "path_to_user_defined_psets_file.json");
                ifcOptions.AddOption("ExportParameterMappingTable", "false"); // Export parameter mapping table
                ifcOptions.AddOption("ExportUserDefinedParameterMappingFile", "path_to_user_defined_parameter_mapping_file.txt");

                // Tab 4: Level of Detail

                // Tab 5: Advanced
                ifcOptions.AddOption("ExportPartsAsBuildingElements", "false");
                ifcOptions.AddOption("AllowMixedSolidModelRepresentation", "false");
                ifcOptions.AddOption("UseActiveViewGeometry", "false");
                ifcOptions.AddOption("UseActiveViewGeometry", "false");
                ifcOptions.AddOption("UseFamilyAndTypeNameForReference", "false"); // Use family and type name for reference
                ifcOptions.AddOption("Use2DRoomBoundariesForRoomVolume", "false"); // Use 2D room boundaries for room volume
                ifcOptions.AddOption("IncludeIFCSiteElevationInTheSiteLocalPlacementOrigin", "false"); // Include IFC site elevation in the site local placement origin
                ifcOptions.AddOption("StoreIFCGUIDInElementParameterAfterExport", "true"); // Store the IFC GUID in an element parameter after export
                ifcOptions.AddOption("ExportBoundingBox", "false");
                ifcOptions.AddOption("Keep Tessellated Geometry As Triangulation", "false");
                ifcOptions.AddOption("UseTypeNameOnlyForIfcType", "false"); // Use type name only for IFC type
                ifcOptions.AddOption("AddLinkingDataForFM", "false");

                // Export to IFC
                doc.Export(ifcFolderPath, ifcFileName, ifcOptions);

                writer.WriteLine(DateTime.Now.ToString() + ": Successfully exported to IFC: " + ifcFilePath);
            }
            catch (Exception ex)
            {
                writer.WriteLine(DateTime.Now.ToString() + ": An error occurred during IFC export for file " + filePath + ": " + ex.Message);
                writer.WriteLine(ex.StackTrace);
            }
        }

        // Class to suppress warnings
        private class WarningSuppressor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                foreach (FailureMessageAccessor failure in failuresAccessor.GetFailureMessages())
                {
                    if (failure.GetSeverity() == FailureSeverity.Warning)
                    {
                        failuresAccessor.DeleteWarning(failure);
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }

        // Class to handle failure processing
        private class FailureProcessor
        {
            private Application _app;
            private FailureHandlingOptions _options;

            public FailureProcessor(Application app, FailureHandlingOptions options)
            {
                _app = app;
                _options = options;
                _app.FailuresProcessing += OnFailuresProcessing;
            }

            private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
            {
                e.SetProcessingOptions(_options);
            }

            ~FailureProcessor()
            {
                _app.FailuresProcessing -= OnFailuresProcessing;
            }
        }
    }
}
