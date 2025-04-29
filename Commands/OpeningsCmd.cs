using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using TeslaRevitTools.ConvoidOpenings;
using TeslaRevitTools.Openings;
using TeslaRevitTools.RvtExtEvents;

namespace TeslaRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpeningsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                DockPanelModel.SwitchDockPanelVisibility(commandData.Application);
                //OpeningsEvent openEvent = new OpeningsEvent();
                //App.Instance.OpeningsViewModel = new OpeningsViewModel();
                //App.Instance.OpeningsViewModel.InitializeViewModel(commandData.Application.ActiveUIDocument.Document);
                //App.Instance.OpeningsViewModel.TheEvent = ExternalEvent.Create(openEvent);
                //App.Instance.OpeningsViewModel.DisplayWindow();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}
