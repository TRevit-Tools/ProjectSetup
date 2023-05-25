using System;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ProjectSetup
{
    [Transaction(TransactionMode.Manual)]
    public class CopyAndSaveAddin : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;

            try
            {
                // Prompt the user to select the source model
                var openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.Filter = "Revit Files (*.rvt)|*.rvt";
                openFileDialog.Title = "Select the source model";

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string sourcePath = openFileDialog.FileName;

                    // Prompt the user to select the target path and file name
                    var saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                    saveFileDialog.Filter = "Revit Files (*.rvt)|*.rvt";
                    saveFileDialog.Title = "Save the model";

                    if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string targetPath = saveFileDialog.FileName;

                        // Convert the file paths to ModelPath
                        ModelPath sourceModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(sourcePath);
                        ModelPath targetModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(targetPath);

                        // Open the source model
                        OpenOptions openOptions = new OpenOptions();
                        openOptions.Detached = true;
                        Document sourceDoc = app.OpenDocumentFile(sourceModelPath, openOptions);

                        // Save the source document as a new central model
                        SaveAsOptions saveAsOptions = new SaveAsOptions();
                        saveAsOptions.OverwriteExistingFile = true;
                        saveAsOptions.Compact = true;
                        saveAsOptions.MaximumBackups = 1;
                        WorksharingSaveAsOptions worksharingOptions = new WorksharingSaveAsOptions();
                        worksharingOptions.SaveAsCentral = true;
                        saveAsOptions.SetWorksharingOptions(worksharingOptions);
                        sourceDoc.SaveAs(targetModelPath, saveAsOptions);

                        // Copy the source model to the target path
                        File.Copy(sourcePath, targetPath, true);

                        TaskDialog.Show("Success", "Model copied and saved successfully!");

                        sourceDoc.Close(false);

                        return Result.Succeeded;
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Cancelled;
        }
    }
}
