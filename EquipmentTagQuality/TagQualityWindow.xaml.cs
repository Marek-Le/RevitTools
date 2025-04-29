using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public partial class TagQualityWindow : Window
    {
        public TagQualityViewModel TagQualityViewModel { get; set; }

        public TagQualityWindow(TagQualityViewModel tagQualityViewModel)
        {
            TagQualityViewModel = tagQualityViewModel;
            DataContext = TagQualityViewModel;
            //SetOwner();
            InitializeComponent();
            LoadingProgressBar.Visibility = System.Windows.Visibility.Hidden;
            ProgressState.Visibility = System.Windows.Visibility.Hidden;
            AddCategories();
        }

        private void SetOwner()
        {
            WindowHandleSearch search = WindowHandleSearch.MainWindowHandle;
            search.SetAsOwner(this);
        }

        public void AddCategories()
        {
            foreach (ElementCategory ec in TagQualityViewModel.TagQualityDataModel.Categories)
            {
                TagQualityViewModel.Categories.Add(ec);
            }
        }

        private void BtnClick_GetElementsByCategory(object sender, RoutedEventArgs e)
        {
            TagQualityViewModel.TagElements.Clear();
            FindElementsBtn.IsEnabled = false;
            BackgroundWorker worker = new BackgroundWorker();
            LoadingProgressBar.Visibility = System.Windows.Visibility.Visible;
            ProgressState.Visibility = System.Windows.Visibility.Visible;
            ProgressState.Text = "searching...";

            TagQualityViewModel.TagQualityDataModel.Worker = worker;
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerAsync();
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            LoadingProgressBar.Value = (double)e.ProgressPercentage;
            ProgressState.Text = e.UserState.ToString();
        }
        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            foreach (var item in TagQualityViewModel.TagQualityDataModel.TagElements) TagQualityViewModel.TagElements.Add(item);
            if(TagQualityViewModel.CustomParameterNames.Count > 0)
            {
                foreach (var parameterName in TagQualityViewModel.CustomParameterNames)
                {
                    foreach (var item in TagQualityViewModel.TagElements)
                    {
                        item.GetCustomParameter(parameterName);
                    }
                }
            }
            FilterItems();
            FindElementsBtn.IsEnabled = true;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            TagQualityViewModel.TagQualityDataModel.TagQualityAction = TagQualityAction.FindElements;
            TagQualityViewModel.TagQualityDataModel.TheEvent.Raise();
            TagQualityViewModel.TagQualityDataModel.SignalEvent.WaitOne();
            TagQualityViewModel.TagQualityDataModel.SignalEvent.Reset();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsInitialized) return;
            CheckBox checkBox = (CheckBox)sender;
            TagElementColumn tagElementColumn = checkBox.DataContext as TagElementColumn;
            var visState = System.Windows.Visibility.Visible;
            SwitchColumnsVisibility(visState, tagElementColumn);

        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!IsInitialized) return;
            CheckBox checkBox = (CheckBox)sender;
            TagElementColumn tagElementColumn = checkBox.DataContext as TagElementColumn;
            var visState = System.Windows.Visibility.Hidden;
            SwitchColumnsVisibility(visState, tagElementColumn);
        }

        private List<string> GetCurrentValidTags()
        {
            return TagQualityViewModel.TagElements.Where(x => x.TagCheckStatus == TagCheckStatus.Tag_OK || x.TagCheckStatus == TagCheckStatus.Tag_Duplicated).Select(x => x.ElementMark.Replace("BER-", "")).ToList();
        }

        private void SwitchColumnsVisibility(System.Windows.Visibility visState, TagElementColumn tagElementColumn)
        {
            switch (tagElementColumn.ColumnName)
            {
                case "Workset":
                    C0.Visibility = visState;
                    break;
                case "Source File":
                    C1.Visibility = visState;
                    break;
                case "Category":
                    C2.Visibility = visState;
                    break;
                case "Family Name":
                    C3.Visibility = visState;
                    break;
                case "Family Type":
                    C4.Visibility = visState;
                    break;
                case "Element Mark":
                    C5.Visibility = visState;
                    break;
                case "ARCU-Tag-Number":
                    C6.Visibility = visState;
                    break;
                case "Equipment Number":
                    C7.Visibility = visState;
                    break;
                case "PGMM-Designation":
                    C8.Visibility = visState;
                    break;
                case "Tag Compare Status":
                    C9.Visibility = visState;
                    break;
                case "Location":
                    C10.Visibility = visState;
                    break;
                case "GB Equipment Category":
                    C11.Visibility = visState;
                    break;
                case "GB System Name":
                    C12.Visibility = visState;
                    break;
                case "GB Equipment Code":
                    C13.Visibility = visState;
                    break;
                case "Tag Check Status":
                    C14.Visibility = visState;
                    break;
                case "Unique Id":
                    C15.Visibility = visState;
                    break;
                case "Revit Id":
                    C16.Visibility = visState;
                    break;
                case "BBoxString":
                    C17.Visibility = visState;
                    break;
                case "Grids":
                    C18.Visibility = visState;
                    break;
                case "EditStatus":
                    C19.Visibility = visState;
                    break;
                default:
                    break;
            }
        }

        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var selectedRow = MainDataGrid.SelectedItem;
            TextBox t = e.EditingElement as TextBox;
            TagElementModel tagElementModel = selectedRow as TagElementModel;
            if(tagElementModel.TagMeetsCriteria(tagElementModel.ElementMark))
            {
                string[] tagCodes = tagElementModel.ElementMark.Replace("BER-", "").Split('-');
                tagElementModel.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                tagElementModel.GetCompareStatus();
                tagElementModel.CheckTag();
                List<string> allTags = GetCurrentValidTags();
                tagElementModel.CheckDuplicates(allTags);
            }
            else
            {
                tagElementModel.GetCompareStatus();
                //e.Cancel = true;             
            }            
        }

        private void FilterBox_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                if (string.IsNullOrEmpty(FilterBox.Text) || FilterBox.Text == "Search filter")
                {
                    FilterBox.Text = "Search filter";
                    FilterBox.Foreground = Brushes.Gray;
                }
                FilterItems();
                MainDataGrid.Focus();
            }
        }

        private void BtnClick_SendSelectedToGb(object sender, RoutedEventArgs e)
        {
            List<TagElementModel> selected = new List<TagElementModel>();
            foreach (TagElementModel item in MainDataGrid.SelectedItems)
            {
                if (item == null) continue;
                selected.Add(item);
            }
            if (selected.Count == 0) return;
            TagQualityViewModel.TagQualityDataModel.SendToGigabase(selected);
        }

        private void BtnClick_SaveAsCsv(object sender, RoutedEventArgs e)
        {
            TagQualityViewModel.TagQualityDataModel.SaveAsCsv();
        }

        private void StatusChecked(object sender, RoutedEventArgs e)
        {
            FilterItems();
        }

        private void FilterItems()
        {
            Predicate<object> filter = new Predicate<object>(item => (item as TagElementModel).EditStatus == "Modified/Not updated" || TagQualityViewModel.PredicateItemStatus(item as TagElementModel) && TagQualityViewModel.PredicateFilterBox(item as TagElementModel, FilterBox.Text));
            MainDataGrid.Items.Filter = filter;
            int itemsShown = MainDataGrid.Items.Count;
            int totalItems = TagQualityViewModel.TagElements.Count;
            SummaryLabel.Content = string.Format("Showing {0} of {1} elements", itemsShown, totalItems);
        }

        private void FilterBox_GotFocus(object sender, RoutedEventArgs e)
        {
            FilterBox.Foreground = Brushes.Black;
            if (FilterBox.Text == "Search filter") FilterBox.Text = "";
        }

        private void CopyArcu_Click(object sender, RoutedEventArgs e)
        {
            foreach (TagElementModel model in MainDataGrid.SelectedItems)
            {
                model.ElementMark = model.ArcuTag;
                if (model.TagMeetsCriteria(model.ElementMark))
                {
                    string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                    model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                    model.GetCompareStatus();
                    model.CheckTag();
                    List<string> allTags = GetCurrentValidTags();
                    model.CheckDuplicates(allTags);
                }
                else
                {
                    model.GetCompareStatus();
                }
            }
        }

        private void CopyPgmm_Click(object sender, RoutedEventArgs e)
        {
            foreach (TagElementModel model in MainDataGrid.SelectedItems)
            {
                model.ElementMark = model.PgmmDesignation;
                if (model.TagMeetsCriteria(model.ElementMark))
                {
                    string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                    model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                    model.GetCompareStatus();
                    model.CheckTag();
                    List<string> allTags = GetCurrentValidTags();
                    model.CheckDuplicates(allTags);
                }
                else
                {
                    model.GetCompareStatus();
                }
            }
        }

        private void CopyEqNumber_Click(object sender, RoutedEventArgs e)
        {
            foreach (TagElementModel model in MainDataGrid.SelectedItems)
            {
                model.ElementMark = model.EquipmentNumber;
                if (model.TagMeetsCriteria(model.ElementMark))
                {
                    string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                    model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                    model.GetCompareStatus();
                    model.CheckTag();
                    List<string> allTags = GetCurrentValidTags();
                    model.CheckDuplicates(allTags);
                }
                else
                {
                    model.GetCompareStatus();
                }
            }
        }

        private void RemoveArcu_Click(object sender, RoutedEventArgs e)
        {
            foreach (TagElementModel model in MainDataGrid.SelectedItems)
            {
                model.ArcuTag = "";
                if (model.TagMeetsCriteria(model.ElementMark))
                {
                    string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                    model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                    model.GetCompareStatus();
                    model.CheckTag();
                }
                else
                {
                    model.GetCompareStatus();
                }
            }
        }

        private void FindReplace_Click(object sender, RoutedEventArgs e)
        {
            FindReplaceWindow window = new FindReplaceWindow(this);
            window.ShowDialog();
            if(window.WindowResult == true)
            {
                foreach (TagElementModel model in MainDataGrid.SelectedItems)
                {
                    model.ElementMark = model.ElementMark.Replace(window.OldBox.Text, window.NewBox.Text);
                    if (model.TagMeetsCriteria(model.ElementMark))
                    {
                        string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                        model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                        model.GetCompareStatus();
                        model.CheckTag();
                        List<string> allTags = GetCurrentValidTags();
                        model.CheckDuplicates(allTags);
                    }
                    else
                    {
                        model.GetCompareStatus();
                    }
                }
            }
        }

        private void SelectInRevit_Click(object sender, RoutedEventArgs e)
        {
            if (MainDataGrid.SelectedItems.Count == 0) return;
            List<Element> elements = new List<Element>();
            foreach (TagElementModel item in MainDataGrid.SelectedItems) elements.Add(item.Element);
            if (elements.Count > 0)
            {
                TagQualityViewModel.TagQualityDataModel.SelectedElements = elements;
                TagQualityViewModel.TagQualityDataModel.TagQualityAction = TagQualityAction.SelectElements;
                TagQualityViewModel.TagQualityDataModel.TheEvent.Raise();
            }
        }

        private void BtnClick_UpdateMark(object sender, RoutedEventArgs e)
        {
            TagQualityViewModel.TagQualityDataModel.TagQualityAction = TagQualityAction.UpdateAllTags;
            TagQualityViewModel.TagQualityDataModel.TheEvent.Raise();
            TagQualityViewModel.TagQualityDataModel.SignalEvent.WaitOne();
            TagQualityViewModel.TagQualityDataModel.SignalEvent.Reset();
            TagQualityViewModel.SummaryLabel = "Update completed!";
            Topmost = true;
            Topmost = false;
        }

        private void ResolveSelectedDuplicates_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<TagElementModel> selectedModels = new List<TagElementModel>();
                foreach (TagElementModel model in MainDataGrid.SelectedItems) selectedModels.Add(model);
                if (selectedModels.Count > 0)
                {
                    var list = selectedModels.Where(x => x.TagCheckStatus == TagCheckStatus.Tag_Duplicated).ToList();
                    if (list.Count != MainDataGrid.SelectedItems.Count)
                    {
                        MessageBox.Show("Selected elements are not duplicates!", "Info");
                        return;
                    }
                    var list2 = selectedModels.Select(x => x.ElementMark.Replace("BER-", "")).Distinct().ToList();
                    if (list2.Count > 1)
                    {
                        MessageBox.Show("Selected elements contain different tags!", "Info");
                        return;
                    }
                }
                TagQualityViewModel.ResolveConflicts(selectedModels);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't resolve conflicts!", "Error");
            }


        }

        private void BtnClick_UpdateSelectedMarks(object sender, RoutedEventArgs e)
        {
            if (MainDataGrid.SelectedItems.Count > 0)
            {
                TagQualityViewModel.TagQualityDataModel.SelectedTagElements = new List<TagElementModel>();
                foreach(TagElementModel model in MainDataGrid.SelectedItems)
                {
                    if (model != null) TagQualityViewModel.TagQualityDataModel.SelectedTagElements.Add(model);
                }
    
                TagQualityViewModel.TagQualityDataModel.TagQualityAction = TagQualityAction.UpdateSelected;
                TagQualityViewModel.TagQualityDataModel.TheEvent.Raise();
                TagQualityViewModel.TagQualityDataModel.SignalEvent.WaitOne();
                TagQualityViewModel.TagQualityDataModel.SignalEvent.Reset();
                TagQualityViewModel.SummaryLabel = "Update completed!";
            }
            Topmost = true;
            Topmost = false;
        }

        private void AddNewParameter_Click(object sender, RoutedEventArgs e)
        {
            string parameterName = "";

            NewParameterWindow newParameterWindow = new NewParameterWindow(this);
            newParameterWindow.ShowDialog();

            if(newParameterWindow.WindowResult)
            {
                parameterName = newParameterWindow.ParameterName.Text;
                if (string.IsNullOrEmpty(parameterName)) return;

                if (TagQualityViewModel.CustomParameterNames.Contains(parameterName))
                {
                    MessageBox.Show("Parameter already added!");
                    return;
                }
                
                TagQualityViewModel.CustomParameterNames.Add(parameterName);

                foreach (var item in TagQualityViewModel.TagElements)
                {
                    item.GetCustomParameter(parameterName);
                }

                DataGridTextColumn dataGridColumn = new DataGridTextColumn();
                dataGridColumn.Header = parameterName;
                dataGridColumn.Binding = new System.Windows.Data.Binding(String.Format("ParameterValues[{0}]", parameterName));
                MainDataGrid.Columns.Add(dataGridColumn);
            }
        }

        private void CopyCustomParameter_Click(object sender, RoutedEventArgs e)
        {
            if(TagQualityViewModel.CustomParameterNames.Count > 0)
            {
                SelectParameterWindow selectParameterWindow = new SelectParameterWindow(TagQualityViewModel.CustomParameterNames, this);
                selectParameterWindow.ShowDialog();
                if(selectParameterWindow.WindowResult)
                {
                    string selectedParameter = selectParameterWindow.ParameterName.SelectedItem as string;
                    if(string.IsNullOrEmpty(selectedParameter)) return;

                    foreach (TagElementModel model in MainDataGrid.SelectedItems)
                    {
                        model.CopyFromCustomParameter(selectedParameter);

                        if (model.TagMeetsCriteria(model.ElementMark))
                        {
                            string[] tagCodes = model.ElementMark.Replace("BER-", "").Split('-');
                            model.ReadTagCodes(tagCodes, TagQualityViewModel.TagQualityDataModel.ExportManager);
                            model.GetCompareStatus();
                            model.CheckTag();
                            List<string> allTags = GetCurrentValidTags();
                            model.CheckDuplicates(allTags);
                        }
                        else
                        {
                            model.GetCompareStatus();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("There is not custom parameters added!");
            }
        }
    }
}
