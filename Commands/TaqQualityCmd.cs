using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TeslaRevitTools.AdvancedElementSelector;
using TeslaRevitTools.EquipmentTagQuality;
using TeslaRevitTools.RvtExtEvents;

namespace TeslaRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TaqQualityCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                TagQualityCheckExEvent tagQualityCheckExEvent = new TagQualityCheckExEvent();             
                TagQualityDataModel dataModel = TagQualityDataModel.Initialize(ExternalEvent.Create(tagQualityCheckExEvent));
                TagQualityViewModel viewModel = new TagQualityViewModel(dataModel);
                App.Instance.TagQualityViewModel = viewModel;
                viewModel.DisplayWindow();
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
