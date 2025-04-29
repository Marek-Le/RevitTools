using Newtonsoft.Json.Bson;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class TagQualityViewModel :INotifyPropertyChanged
    {
        public List<string> CustomParameterNames { get; set; }
        private string _summaryLabel;
        public string SummaryLabel
        {
            get => _summaryLabel;
            set
            {
                if (_summaryLabel != value)
                {
                    _summaryLabel = value;
                    OnPropertyChanged(nameof(SummaryLabel));
                }
            }
        }
        public ObservableCollection<CompareStatusCheck> StatusChecks { get; set; }
        public ObservableCollection<ElementCategory> Categories { get; set; }
        public ObservableCollection<TagElementModel> TagElements { get; set; }
        public ObservableCollection<TagElementColumn> ColumnsVisibility { get; set; }
        public bool ActiveViewOnly { get; set; } = false;
        public bool ExcludeSubcomponents { get; set; } = true;

        public TagQualityDataModel TagQualityDataModel { get; set; }
        //public Thread WindowThread { get; set; }
        
        public TagQualityViewModel(TagQualityDataModel tagQualityDataModel)
        {
            TagQualityDataModel = tagQualityDataModel;
            Categories = new ObservableCollection<ElementCategory>();
            TagElements = new ObservableCollection<TagElementModel>();
            SetColumnsVisibility();
            SetStatusChecks();
            CustomParameterNames = new List<string>();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void ResolveConflicts(List<TagElementModel> selectedModels)
        {

            List<int> codes = TagElements.Where(x => x.TagCheckStatus == TagCheckStatus.Tag_OK || x.TagCheckStatus == TagCheckStatus.Tag_Duplicated)
                                .Where(x => x.Location == selectedModels.First().Location && x.EquipmentCategory == selectedModels.First().EquipmentCategory && x.SystemName == selectedModels.First().SystemName)
                                    .Select(x => int.Parse(x.ElementMark.Replace("BER-", "").Split('-')[3])).ToList();
            selectedModels = selectedModels.OrderBy(x => x.ElementMark).ToList();
            int equipmentCode = int.Parse(selectedModels[0].ElementMark.Replace("BER-", "").Split('-')[3]) + 1;

            for (int i = 1; i < selectedModels.Count; i++)
            {
                while (codes.Contains(equipmentCode)) equipmentCode++;
                codes.Add(equipmentCode);
                selectedModels[i].ElementMark = "BER-" + selectedModels[i].ElementMark.Replace("BER-", "").Split('-')[0] + "-" + selectedModels[i].ElementMark.Replace("BER-", "").Split('-')[1] + "-" + selectedModels[i].ElementMark.Replace("BER-", "").Split('-')[2] + "-" + equipmentCode.ToString("D3");
            }
        }

        public bool PredicateFilterBox(TagElementModel tagElementModel, string filterText)
        {
            if (string.IsNullOrEmpty(filterText) || filterText == "Search filter") return true;
            if (tagElementModel.GetHash().ToUpper().Contains(filterText.ToUpper())) return true;
            return false;
        }

        public bool PredicateItemStatus(TagElementModel tagElementModel)
        {
            bool result = false;
            List<CompareStatusCheck> compareStatuses = StatusChecks.Where(e => e.IsChecked).ToList();
            foreach (CompareStatusCheck compareStatusCheck in compareStatuses)
            {
                if (tagElementModel.CompareStatus == compareStatusCheck.Status) result = true;
            }
            return result;
        }

        public void DisplayWindow()
        {
            Thread windowThread = new Thread(delegate ()
            {
                TagQualityWindow window = new TagQualityWindow(this);
                window.Show();
                Dispatcher.Run();
            });
            windowThread.SetApartmentState(ApartmentState.STA);
            windowThread.IsBackground = true;
            windowThread.Start();
            //WindowThread = windowThread;
        }

        private void SetColumnsVisibility()
        {
            ColumnsVisibility = new ObservableCollection<TagElementColumn>()
            {
                new TagElementColumn() { ColumnName = "Workset", IsChecked = true },
                new TagElementColumn() { ColumnName = "Source File", IsChecked = false },
                new TagElementColumn() { ColumnName = "Category", IsChecked = true },
                new TagElementColumn() { ColumnName = "Family Name", IsChecked = true },
                new TagElementColumn() { ColumnName = "Family Type", IsChecked = true },
                new TagElementColumn() { ColumnName = "Element Mark", IsChecked = true },
                new TagElementColumn() { ColumnName = "ARCU-Tag-Number", IsChecked = true },
                new TagElementColumn() { ColumnName = "Equipment Number", IsChecked = false },
                new TagElementColumn() { ColumnName = "PGMM-Designation", IsChecked = false },
                new TagElementColumn() { ColumnName = "Tag Compare Status", IsChecked = true },
                new TagElementColumn() { ColumnName = "Location", IsChecked = true },
                new TagElementColumn() { ColumnName = "GB Equipment Category", IsChecked = true },
                new TagElementColumn() { ColumnName = "GB System Name", IsChecked = true },
                new TagElementColumn() { ColumnName = "GB Equipment Code", IsChecked = true },
                new TagElementColumn() { ColumnName = "Tag Check Status", IsChecked = true },
                new TagElementColumn() { ColumnName = "Unique Id", IsChecked = false },
                new TagElementColumn() { ColumnName = "Revit Id", IsChecked = true },
                new TagElementColumn() { ColumnName = "BBoxString", IsChecked = false },
                new TagElementColumn() { ColumnName = "Grids", IsChecked = false },
                new TagElementColumn() { ColumnName = "EditStatus", IsChecked = true },
            };
        }

        private void SetStatusChecks()
        {
            StatusChecks = new ObservableCollection<CompareStatusCheck>()
            {
                new CompareStatusCheck() { IsChecked = true, Status = CompareStatus.Status1 },
                new CompareStatusCheck() { IsChecked = true, Status = CompareStatus.Status2 },
                new CompareStatusCheck() { IsChecked = true, Status = CompareStatus.Status3 },
                new CompareStatusCheck() { IsChecked = true, Status = CompareStatus.Status4 },
                new CompareStatusCheck() { IsChecked = true, Status = CompareStatus.Status5 }
            };
        }
    }
}
