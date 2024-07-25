using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using System.Windows.Forms;

namespace ClassLibrary2
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string folderPath = SelectFolder();
            if (string.IsNullOrEmpty(folderPath))
            {
                message = "No folder selected.";
                return Result.Failed;
            }

            string logFilePath = Path.Combine(folderPath, "Dynamo_Revit_Log.txt");

            try
            {
                UIApplication uiApp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;

                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    foreach (string filePath in files)
                    {
                        Document doc = null;
                        bool saveChanges = false;

                        try
                        {
                            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                            OpenOptions openOptions = new OpenOptions();
                            doc = app.OpenDocumentFile(modelPath, openOptions);

                            if (doc != null)
                            {
                                writer.WriteLine(DateTime.Now.ToString() + ": Successfully opened Revit file: " + filePath);

                                bool mismatchFound = CheckPanelCircuitMismatch(doc, writer);

                                if (mismatchFound)
                                {
                                    using (Transaction trans = new Transaction(doc, "Disconnect Panel from Circuit"))
                                    {
                                        trans.Start();
                                        try
                                        {
                                            DisconnectPanelFromCircuit(doc, writer);
                                            saveChanges = true;
                                            writer.WriteLine(DateTime.Now.ToString() + ": Disconnected panel from circuit for file: " + filePath);
                                            trans.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                            trans.RollBack();
                                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred during the transaction for file " + filePath + ": " + ex.Message);
                                            writer.WriteLine(ex.StackTrace);
                                        }
                                    }
                                }
                                else
                                {
                                    saveChanges = true;
                                }

                                if (saveChanges)
                                {
                                    SaveDocument(doc, writer);
                                }

                                ExportToIFC(doc, filePath, writer);
                            }
                            else
                            {
                                HandleFileOpeningError(filePath, writer);
                            }
                        }
                        catch (Exception ex)
                        {
                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred with file " + filePath + ": " + ex.Message);
                            writer.WriteLine(ex.StackTrace);

                            if (doc != null)
                            {
                                try
                                {
                                    using (Transaction trans = new Transaction(doc, "Disconnect Panel from Circuit on Error"))
                                    {
                                        trans.Start();
                                        DisconnectPanelFromCircuit(doc, writer);
                                        saveChanges = true;
                                        writer.WriteLine(DateTime.Now.ToString() + ": Automatically disconnected panel from circuit for file: " + filePath);
                                        trans.Commit();
                                    }
                                }
                                catch (Exception innerEx)
                                {
                                    writer.WriteLine(DateTime.Now.ToString() + ": An additional error occurred while trying to disconnect the panel for file " + filePath + ": " + innerEx.Message);
                                    writer.WriteLine(innerEx.StackTrace);
                                }
                            }
                        }
                        finally
                        {
                            if (doc != null && doc.IsValidObject)
                            {
                                doc.Close(false);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
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

        private string SelectFolder()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                DialogResult result = folderDialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                {
                    return folderDialog.SelectedPath;
                }
            }
            return string.Empty;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer)
        {
            writer.WriteLine($"{DateTime.Now}: Failed to open Revit file: {filePath}");
        }

        private void SaveDocument(Document doc, StreamWriter writer)
        {
            try
            {
                doc.Save();
                writer.WriteLine($"{DateTime.Now}: Document saved successfully.");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{DateTime.Now}: An error occurred while saving the document: {ex.Message}");
                writer.WriteLine(ex.StackTrace);
            }
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

        private void DisconnectPanelFromCircuit(Document doc, StreamWriter writer)
        {
            var panels = new FilteredElementCollector(doc)
                         .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                         .OfClass(typeof(FamilyInstance))
                         .WhereElementIsNotElementType()
                         .OfType<FamilyInstance>()
                         .ToList();

            foreach (var panel in panels)
            {
                var electricalSystems = panel.MEPModel?.GetElectricalSystems();
                if (electricalSystems != null)
                {
                    foreach (var electricalSystem in electricalSystems)
                    {
                        using (var trans = new Transaction(doc, "Disconnect Panel from Circuit"))
                        {
                            trans.Start();
                            foreach (var connector in electricalSystem.ConnectorManager.Connectors.OfType<Connector>())
                            {
                                foreach (var connectedConnector in connector.AllRefs.OfType<Connector>())
                                {
                                    connector.DisconnectFrom(connectedConnector);
                                    writer.WriteLine($"{DateTime.Now}: Disconnected connector {connector.Owner.Id} from {connectedConnector.Owner.Id}");
                                }
                            }
                            if (ShouldDeleteElement(electricalSystem))
                            {
                                doc.Delete(electricalSystem.Id);
                                writer.WriteLine($"{DateTime.Now}: Deleted electrical system: {electricalSystem.Id.IntegerValue}");
                            }
                            trans.Commit();
                        }
                    }
                }
            }
        }

        private bool ShouldDeleteElement(Element element)
        {
            if (element is ElectricalSystem electricalSystem)
            {
                List<Connector> loadConnectors = new List<Connector>();

                foreach (Element e in electricalSystem.Elements)
                {
                    if (e is FamilyInstance familyInstance && familyInstance.MEPModel != null)
                    {
                        var connectors = familyInstance.MEPModel.ConnectorManager.Connectors
                            .OfType<Connector>()
                            .Where(c => c.IsConnected);

                        loadConnectors.AddRange(connectors);
                    }
                }

                return !loadConnectors.Any();
            }

            return false;
        }

        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                string ifcFolderPath = Path.Combine(Path.GetDirectoryName(filePath), "IFC_Exports");
                Directory.CreateDirectory(ifcFolderPath);

                string ifcFileName = Path.GetFileNameWithoutExtension(filePath) + ".ifc";
                string ifcFilePath = Path.Combine(ifcFolderPath, ifcFileName);

                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    FileVersion = IFCVersion.IFC2x3CV2,
                    ExportBaseQuantities = true,
                    SpaceBoundaryLevel = 2
                };

                ifcOptions.AddOption("ExchangeRequirement", "IFC4_Reference_View");
                ifcOptions.AddOption("FileType", "IFC2x3");

                PhaseArray phases = doc.Phases;
                Phase phaseToExport = phases.get_Item(phases.Size - 1);
                ifcOptions.AddOption("ExportingPhase", phaseToExport.Id.IntegerValue.ToString());
                ifcOptions.AddOption("AdditionalProximityControl", "true");

                ifcOptions.AddOption("ExportLinkedFiles", "false");
                ifcOptions.AddOption("ExportRevitPropertySets", "false");
                ifcOptions.AddOption("ExportIFCCommonPropertySets", "true");
                ifcOptions.ExportBaseQuantities = false;
                ifcOptions.AddOption("ExportMaterialPropertySets", "false");
                ifcOptions.AddOption("ExportSchedulesAsPsets", "false");
                ifcOptions.AddOption("ExportOnlySchedulesContainingIFCPsetOrCommonInTitle", "false");
                ifcOptions.AddOption("ExportUserDefinedPsets", "false");
                ifcOptions.AddOption("ExportUserDefinedPsetsFile", "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\JoaquinDefinedPropertySet.txt");
                ifcOptions.AddOption("ExportParameterMappingTable", "false");
                ifcOptions.AddOption("ExportUserDefinedParameterMappingFile", "path_to_user_defined_parameter_mapping_file.txt");

                ifcOptions.AddOption("ExportPartsAsBuildingElements", "false");
                ifcOptions.AddOption("AllowMixedSolidModelRepresentation", "false");
                ifcOptions.AddOption("UseActiveViewGeometry", "false");
                ifcOptions.AddOption("UseFamilyAndTypeNameForReference", "false");
                ifcOptions.AddOption("Use2DRoomBoundariesForRoomVolume", "false");
                ifcOptions.AddOption("IncludeIFCSiteElevationInTheSiteLocalPlacementOrigin", "false");
                ifcOptions.AddOption("StoreIFCGUIDInElementParameterAfterExport", "true");
                ifcOptions.AddOption("ExportBoundingBox", "false");
                ifcOptions.AddOption("KeepTessellatedGeometryAsTriangulation", "false");
                ifcOptions.AddOption("UseTypeNameOnlyForIfcType", "false");
                ifcOptions.AddOption("UseVisibleRevitNameAsIfcEntityName", "true");

                ifcOptions.AddOption("SitePlacement", "DefaultSite");
                ifcOptions.AddOption("CoordinateBase", "SharedCoordinates");

                using (Transaction exportTrans = new Transaction(doc, "Export IFC"))
                {
                    exportTrans.Start();
                    doc.Export(ifcFolderPath, ifcFileName, ifcOptions);
                    exportTrans.Commit();
                }

                writer.WriteLine($"{DateTime.Now}: IFC export completed for: {ifcFilePath}");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{DateTime.Now}: An error occurred during IFC export for file {filePath}: {ex.Message}");
                writer.WriteLine(ex.StackTrace);
            }
        }
    }
}
