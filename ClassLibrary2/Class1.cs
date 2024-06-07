using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Events;

namespace ClassLibrary2
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string folderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Revit Files";
            string logFilePath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Dynamo_Revit_Log.txt";

            try
            {
                UIApplication uiApp = commandData.Application;
                Application app = uiApp.Application;

                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    // Set up failure message handler
                    app.FailuresProcessing += OnFailuresProcessing;

                    foreach (string filePath in files)
                    {
                        Document doc = null;

                        try
                        {
                            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                            OpenOptions openOptions = new OpenOptions();
                            doc = app.OpenDocumentFile(modelPath, openOptions);

                            if (doc != null)
                            {
                                writer.WriteLine($"{DateTime.Now}: Successfully opened Revit file: {filePath}");

                                // Perform operations on the document
                                DisconnectAllElements(doc, writer);
                                CheckPanelCircuitMismatch(doc, writer);
                                SaveDocument(doc, writer);
                                ExportToIFC(doc, filePath, writer);
                            }
                            else
                            {
                                HandleFileOpeningError(filePath, writer);
                            }
                        }
                        catch (Exception ex)
                        {
                            writer.WriteLine($"{DateTime.Now}: An error occurred with file {filePath}: {ex.Message}");
                            writer.WriteLine(ex.StackTrace);
                        }
                        finally
                        {
                            if (doc != null && doc.IsValidObject)
                            {
                                doc.Close(false);
                            }
                        }
                    }

                    // Remove failure message handler
                    app.FailuresProcessing -= OnFailuresProcessing;
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter writer2 = new StreamWriter(logFilePath, true))
                {
                    writer2.WriteLine($"{DateTime.Now}: An error occurred: {ex.Message}");
                }
            }

            return Result.Succeeded;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer)
        {
            writer.WriteLine($"{DateTime.Now}: Failed to open Revit file: {filePath}");
        }

        private void CheckPanelCircuitMismatch(Document doc, StreamWriter writer)
        {
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
                    writer.WriteLine($"{DateTime.Now}: Mismatch found for panel {panel.Name}. Total load exceeds capacity.");
                }
            }
        }

        private void DisconnectAllElements(Document doc, StreamWriter writer)
        {
            DisconnectElectricalEquipment(doc, writer);
            DisconnectPipes(doc, writer);
            DisconnectDucts(doc, writer);
            DisconnectCircuits(doc, writer);
        }

        private void DisconnectElectricalEquipment(Document doc, StreamWriter writer)
        {
            var electricalElements = new FilteredElementCollector(doc)
                                     .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                                     .OfClass(typeof(FamilyInstance))
                                     .WhereElementIsNotElementType()
                                     .OfType<FamilyInstance>()
                                     .ToList();

            foreach (var element in electricalElements)
            {
                FamilyInstance familyInstance = element as FamilyInstance;
                if (familyInstance != null)
                {
                    DisconnectConnectors(familyInstance.MEPModel.ConnectorManager, doc, writer, "Disconnect Electrical Equipment");
                }
            }
        }

        private void DisconnectPipes(Document doc, StreamWriter writer)
        {
            var pipes = new FilteredElementCollector(doc)
                        .OfClass(typeof(Pipe))
                        .WhereElementIsNotElementType()
                        .ToList();

            foreach (var pipe in pipes)
            {
                DisconnectConnectors(((Pipe)pipe).ConnectorManager, doc, writer, "Disconnect Pipe");
            }
        }

        private void DisconnectDucts(Document doc, StreamWriter writer)
        {
            var ducts = new FilteredElementCollector(doc)
                        .OfClass(typeof(Duct))
                        .WhereElementIsNotElementType()
                        .ToList();

            foreach (var duct in ducts)
            {
                DisconnectConnectors(((Duct)
                    duct).ConnectorManager, doc, writer, "Disconnect Duct");
            }
        }

        private void DisconnectCircuits(Document doc, StreamWriter writer)
        {
            var circuits = new FilteredElementCollector(doc)
                            .OfClass(typeof(ElectricalSystem))
                            .WhereElementIsNotElementType()
                            .Cast<ElectricalSystem>()
                            .ToList();

            foreach (var circuit in circuits)
            {
                DisconnectConnectors(circuit.ConnectorManager, doc, writer, "Disconnect Circuit");
            }
        }

        private void DisconnectConnectors(ConnectorManager connectorManager, Document doc, StreamWriter writer, string transactionName)
        {
            using (var trans = new Transaction(doc, transactionName))
            {
                trans.Start();

                var connectors = connectorManager.Connectors.Cast<Connector>();

                foreach (var connector in connectors)
                {
                    if (connector.IsConnected)
                    {
                        var connectedConnectors = connector.AllRefs.OfType<Connector>().ToList();

                        foreach (var connectedConnector in connectedConnectors)
                        {
                            try
                            {
                                // Disconnect the main connector from each connected connector
                                connector.DisconnectFrom(connectedConnector);
                            }
                            catch (Exception ex)
                            {
                                writer.WriteLine($"{DateTime.Now}: An error occurred while disconnecting: {ex.Message}");
                                writer.WriteLine(ex.StackTrace);
                            }
                        }
                    }
                }

                trans.Commit();
            }
        }

        private void SaveDocument(Document doc, StreamWriter writer)
        {
            try
            {
                using (Transaction trans = new Transaction(doc, "Save Document"))
                {
                    trans.Start();
                    doc.Save();
                    trans.Commit();
                }
                writer.WriteLine($"{DateTime.Now}: Document saved successfully.");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{DateTime.Now}: An error occurred while saving the document: {ex.Message}");
                writer.WriteLine(ex.StackTrace);
            }
        }

        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    FileVersion = IFCVersion.IFC2x3CV2,
                    ExportBaseQuantities = true
                };
                string ifcFolderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\R2024";
                Directory.CreateDirectory(ifcFolderPath);

                string ifcFileName = Path.GetFileNameWithoutExtension(filePath) + ".ifc";
                string ifcFilePath = Path.Combine(ifcFolderPath, ifcFileName);

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

        // Failure processing handler to suppress errors and warnings
        private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            var failList = e.GetFailuresAccessor().GetFailureMessages();
            foreach (var failureMessageAccessor in failList)
            {
                // Dismiss all types of warnings and failures
                e.GetFailuresAccessor().DeleteWarning(failureMessageAccessor);
            }
            e.SetProcessingResult(FailureProcessingResult.Continue);
        }
    }
}