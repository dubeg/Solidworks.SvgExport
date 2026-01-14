using CADBooster.SolidDna;
using Dx.Sw.SvgExport.Command;
using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;
using static CADBooster.SolidDna.SolidWorksEnvironment;

namespace Dx.Sw.SvgExport;

[Guid("ca536ccd-e9d9-4f9a-a74b-3732f1bcb08c"), ComVisible(true)]
public class SvgPlugin : SolidPlugIn<SvgPlugin> {
    public override string AddInTitle => "SVG Export Plugin";
    public override string AddInDescription => "Exports SVG(s) from slddrw views";
    public override void ConnectedToSolidWorks() {
        var commandManagerItems = new List<CommandManagerItem> {
            new CommandManagerItem {
                Name = "Export to SVG",
                Tooltip = "Export the current drawing view to an SVG file",
                ImageIndex = 0,
                Hint = "Export the current drawing view to an SVG file",
                VisibleForDrawings = true,
                VisibleForAssemblies = false,
                VisibleForParts = false,
                OnClick = () => {
                    try {
                        var cmd = new DrawingToSvgCommand(Application);
                        cmd.RunForCurrentSheet();
                    }
                    catch(Exception ex) { 
                        Application.ShowMessageBox(ex.Message);
                    }
                },
                OnStateCheck = (args) => args.Result = CommandManagerItemState.DeselectedEnabled
            }
        };

        Application.CommandManager.CreateCommandTab(
            title: "SVG",
            id: 150_001,
            commandManagerItems: commandManagerItems.Select(x => (ICommandManagerItem)x).ToList()
        );
    }
    public override void DisconnectedFromSolidWorks() {}
}

