using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CADBooster.SolidDna;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Dx.Sw.SvgExport.Command;

public class DrawingToSvgCommand(SolidWorksApplication App) {
    
    private string PromptForOutputFolder() {
        var folderBrowserDialog = new FolderBrowserEx.FolderBrowserDialog {
            Title = "Select output folder for exported SVG files",
            AllowMultiSelect = false,
        };
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) {
            return folderBrowserDialog.SelectedFolder;
        }
        return null;
    }

    public string RunForCurrentSheet() {
        ModelDoc2 model = null;
        try {
            var outputFolderPath = PromptForOutputFolder();
            if (string.IsNullOrEmpty(outputFolderPath)) {
                return null;
            }
            if (App.ActiveModel == null) {
                throw new InvalidOperationException("No active document.");
            }
            model = App.ActiveModel.UnsafeObject;
            string selectedViewName = null;
            var selMgr = model.ISelectionManager;
            if (selMgr.GetSelectedObjectCount2(-1) == 1) {
                var selType = (swSelectType_e)selMgr.GetSelectedObjectType3(1, -1);
                if (selType == swSelectType_e.swSelDRAWINGVIEWS) {
                    var selectedView = (IView)selMgr.GetSelectedObject6(1, -1);
                    selectedViewName = selectedView.GetName2();
                }
            }
            model.FeatureManager.EnableFeatureTree = false;
            model.FeatureManager.EnableFeatureTreeWindow = false;
            model.ConfigurationManager.EnableConfigurationTree = false;
            model.IActiveView.EnableGraphicsUpdate = false;
            model.SketchManager.DisplayWhenAdded = false;
            model.SketchManager.AddToDB = true;
            model.ISelectionManager.EnableContourSelection = false;
            model.ISelectionManager.EnableSelection = false;
            model.Lock();
            if (!Directory.Exists(outputFolderPath)) {
                Directory.CreateDirectory(outputFolderPath);
            }
            var baseFileName = Path.GetFileNameWithoutExtension(model.GetPathName());
            var fileName = string.IsNullOrEmpty(selectedViewName) 
                ? $"{baseFileName}.svg" 
                : $"{baseFileName}_{selectedViewName}.svg";
            var outFilePath = Path.Combine(outputFolderPath, fileName);
            var exporter = new SvgExporter(App.UnsafeObject);
            var exportIndividualViews = string.IsNullOrEmpty(selectedViewName); // Only export individual views if no specific view selected
            var warnings = exporter.Export(outFilePath, fitToContent: true, includeBomMetadata: true, exportIndividualViews: exportIndividualViews, specificViewName: selectedViewName);
            System.Diagnostics.Process.Start("explorer", $""" "{outFilePath}" """);
            return outFilePath;
        }
        finally {
            if (model is not null) {
                model.UnLock();
                model.FeatureManager.EnableFeatureTree = true;
                model.FeatureManager.EnableFeatureTreeWindow = true;
                model.ConfigurationManager.EnableConfigurationTree = true;
                model.IActiveView.EnableGraphicsUpdate = true;
                model.SketchManager.DisplayWhenAdded = true;
                model.SketchManager.AddToDB = false;
                model.FeatureManager.UpdateFeatureTree();
                model.ISelectionManager.EnableContourSelection = false;
                model.ISelectionManager.EnableSelection = true;
            }
        }
    }
}

