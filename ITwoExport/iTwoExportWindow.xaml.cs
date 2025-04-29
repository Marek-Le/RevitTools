using RevitUI = Autodesk.Revit.UI;
using CefSharp;
using DocumentFormat.OpenXml.Drawing;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;

namespace TeslaRevitTools.ITwoExport
{
    /// <summary>
    /// Interaction logic for iTwoExportWindow.xaml
    /// </summary>
    public partial class iTwoExportWindow : Window
    {
        public ExportViewModel ExportViewModel { get; set; }
        public iTwoExportWindow(ExportViewModel exportViewModel, IntPtr owner)
        {
            SetAsOwner(this, owner);
            ExportViewModel = exportViewModel;
            DataContext = ExportViewModel;
            InitializeComponent();
            Loaded += ITwoExportWindow_Initialized;
        }

        private void ITwoExportWindow_Initialized(object sender, EventArgs e)
        {
            HeaderCheckbox.IsChecked = GetHeaderCheckboxState();

            //align column order
            foreach (PreviewColumn pc in ExportViewModel.ElementsPreviewColumns)
            {
                DataGridColumn column = PreviewDataGrid.Columns.Where(x => x.Header.ToString() == pc.HeaderName).FirstOrDefault();
                if(!pc.IsChecked)
                {
                    column.Visibility = Visibility.Collapsed;
                }
                else
                {
                    column.Visibility = Visibility.Visible;
                }
                column.DisplayIndex = pc.OrderIndex;
            }

            foreach (PreviewColumn pc in ExportViewModel.MaterialsPreviewColumns)
            {
                DataGridColumn column = MaterialsPreviewDataGrid.Columns.Where(x => x.Header.ToString() == pc.HeaderName).FirstOrDefault();
                if (!pc.IsChecked)
                {
                    column.Visibility = Visibility.Collapsed;
                }
                else
                {
                    column.Visibility = Visibility.Visible;
                }
                column.DisplayIndex = pc.OrderIndex;
            }

        }

        private void SetAsOwner(Window childWindow, IntPtr handle)
        {
            var helper = new WindowInteropHelper(childWindow) { Owner = handle };
        }

        private void BtnClick_AppendExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportViewModel.GetExcelData(false);
                if (ExportViewModel.ExcelKeynotes.Count > 0) ExportViewModel.CheckElementKeynotes();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error", "Expected .xlsx or .xlsm format, sheet name: Keynotes" + Environment.NewLine + ex.ToString());
            }
        }

        private void BtnClick_LoadExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportViewModel.GetExcelData(true);
                if (ExportViewModel.ExcelKeynotes.Count > 0) ExportViewModel.CheckElementKeynotes();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error", "Expected .xlsx or .xlsm format, sheet name: Keynotes" + Environment.NewLine + ex.ToString());
            }
        }

        private void MenuItemClick_SetPreview(object sender, RoutedEventArgs e)
        {
            ExportViewModel.SetPreviewItems();
            ElementsSearchBox.Text = "";
            MainDataGrid.Items.Filter = null;
        }

        private void MainDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            TextBox t = e.EditingElement as TextBox;
            string editedCellValue = t.Text.ToString();
            if (originalKeynote == null)
            {
                return;
            }
            if (editedCellValue != originalKeynote)
            {
                foreach (ExportedElement ee in ExportViewModel.ExportedElements)
                {
                    if (ee.Keynote == originalKeynote)
                    {
                        ee.Keynote = editedCellValue;
                        ee.TrySplitKeynote();
                    }
                }
            }
        }

        private string originalKeynote;

        private void MainDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            ExportedElement activeCellAtEdit = MainDataGrid.SelectedItem as ExportedElement;
            if(activeCellAtEdit == null) return;
            if (string.IsNullOrEmpty(activeCellAtEdit.Keynote))
            {
                originalKeynote = null;
                return;
            }
            originalKeynote = activeCellAtEdit.Keynote.ToString();
        }

        private void BtnClick_SaveExcelData(object sender, RoutedEventArgs e)
        {
            ExportViewModel.SaveExcelDataAsJson();
        }

        private void MenuItemClick_SaveExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportViewModel.SaveElementsToExcel();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error", ex.ToString());
            }
        }

        private void MenuItemClick_RefreshDataContext(object sender, RoutedEventArgs e)
        {
            ExportViewModel.Action = ITwoAction.GetElements;
            ExportViewModel.TheEvent.Raise();
        }

        private void MenuItemClick_CheckKeynotes(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                ContextMenu contextMenu = menuItem.Parent as ContextMenu;
                if(contextMenu != null)
                {
                    DataGrid dataGrid = contextMenu.PlacementTarget as DataGrid;
                    if(dataGrid != null)
                    {
                        if (dataGrid.Name == "MainDataGrid")
                        {
                            ExportViewModel.CheckElementKeynotes();
                        }
                        if (dataGrid.Name == "MaterialsDataGrid")
                        {
                            ExportViewModel.CheckMaterialKeynotes();
                        }
                        dataGrid.Items.Refresh();
                    }
                }
            }
        }

        private void MenuItemClick_SelectInModel(object sender, RoutedEventArgs e)
        {
            ExportedElement exportedElement = MainDataGrid.SelectedItem as ExportedElement;
            if (exportedElement != null)
            {
                ExportViewModel.SelectedElement = exportedElement;
                ExportViewModel.Action = ITwoAction.SelectInModel;
                ExportViewModel.TheEvent.Raise();
            }
        }

        private void BtnClick_SaveSelectedCategories(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportViewModel.SaveSelectedCategories();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error", ex.ToString());
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var item in ExportViewModel.RevitCategories)
            {
                item.IsSelected = true;
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var item in ExportViewModel.RevitCategories)
            {
                item.IsSelected = false;
            }
        }

        private Nullable<bool> GetHeaderCheckboxState()
        {
            if (!IsInitialized) return null;
            if (ExportViewModel.RevitCategories.All(e => e.IsSelected)) return true;
            if (ExportViewModel.RevitCategories.All(e => !e.IsSelected)) return false;
            return null;
        }

        private bool _allowChckboxEvent = true;

        private void CheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _allowChckboxEvent = true;
        }

        private void CheckBox_Checked_1(object sender, RoutedEventArgs e)
        {
            if (_allowChckboxEvent)
            {
                HeaderCheckbox.IsChecked = GetHeaderCheckboxState();
                _allowChckboxEvent = false;
            }
        }

        private void CheckBox_Unchecked_1(object sender, RoutedEventArgs e)
        {
            if (_allowChckboxEvent)
            {
                HeaderCheckbox.IsChecked = GetHeaderCheckboxState();
                _allowChckboxEvent = false;
            }
        }

        private void MenuItemClick_LoadMaterials(object sender, RoutedEventArgs e)
        {
            ExportViewModel.Action = ITwoAction.GetMaterials;
            ExportViewModel.TheEvent.Raise();
        }

        private void MenuItemClick_SelectElementContaining(object sender, RoutedEventArgs e)
        {
            ExportedMaterial exportedMaterial = MaterialsDataGrid.SelectedItem as ExportedMaterial;
            if (exportedMaterial != null)
            {
                ExportViewModel.SelectedMaterial = exportedMaterial;
                ExportViewModel.Action = ITwoAction.SelectElementContaining;
                ExportViewModel.TheEvent.Raise();
            }
        }

        private void MenuItemClick_SelectMaterial(object sender, RoutedEventArgs e)
        {
            ExportedMaterial exportedMaterial = MaterialsDataGrid.SelectedItem as ExportedMaterial;
            if (exportedMaterial != null)
            {
                ExportViewModel.SelectedMaterial = exportedMaterial;
                ExportViewModel.Action = ITwoAction.SelectMaterial;
                ExportViewModel.TheEvent.Raise();
            }
        }

        private void MenuItemClick_GenerateMaterialsPreview(object sender, RoutedEventArgs e)
        {
            ExportViewModel.SetPreviewMaterials();
            MaterialsSearchBox.Text = "";
            MaterialsDataGrid.Items.Filter = null;
        }

        private void MaterialsSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsInitialized) return;
            if (string.IsNullOrEmpty(MaterialsSearchBox.Text))
            {
                MaterialsDataGrid.Items.Filter = null;
                return;
            }

            Predicate<object> filter = new Predicate<object>(item => (item as ExportedMaterial).Hash.ToUpper().Contains(MaterialsSearchBox.Text.ToUpper()));
            MaterialsDataGrid.Items.Filter = filter;
        }

        private void ElementsSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsInitialized) return;
            if (string.IsNullOrEmpty(ElementsSearchBox.Text))
            {
                MainDataGrid.Items.Filter = null;
                return;
            }

            Predicate<object> filter = new Predicate<object>(item => (item as ExportedElement).Hash.ToUpper().Contains(ElementsSearchBox.Text.ToUpper()));
            MainDataGrid.Items.Filter = filter;

        }

        private void MenuItemClick_SaveMaterialsToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportViewModel.SaveMaterialsToExcel();
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error", ex.ToString());
            }
        }

        private void MenuItem_SetDefaultElementExport(object sender, RoutedEventArgs e)
        {
            var result = RevitUI.TaskDialog.Show("Warning", "You will loose your current config. Are you sure?", RevitUI.TaskDialogCommonButtons.No | RevitUI.TaskDialogCommonButtons.Yes);
            if (result != RevitUI.TaskDialogResult.Yes) { return; }
            var previewColumns = new List<PreviewColumn>
                {
                    new PreviewColumn() { HeaderName = "Type Name", IsChecked = true, BindingParameter = "TypeName", OrderIndex = 0},
                    new PreviewColumn() { HeaderName = "Description", IsChecked = true, BindingParameter = "Description", OrderIndex = 1  },
                    new PreviewColumn() { HeaderName = "KeyNotes", IsChecked = true, BindingParameter = "Keynote", OrderIndex = 2  },
                    new PreviewColumn() { HeaderName = "DIN", IsChecked = true, BindingParameter = "Din", OrderIndex = 3  },
                    new PreviewColumn() { HeaderName = "RN", IsChecked = true, BindingParameter = "Rn", OrderIndex = 4  },
                    new PreviewColumn() { HeaderName = "Short Info", IsChecked = true, BindingParameter = "ShortInfo", OrderIndex = 5  },
                    new PreviewColumn() { HeaderName = "Outl. Spec.", IsChecked = true, BindingParameter = "OutlineSpec", OrderIndex = 6 },
                    new PreviewColumn() { HeaderName = "Quantity", IsChecked = true, BindingParameter = "Quantity" , OrderIndex = 7 },
                    new PreviewColumn() { HeaderName = "UoM", IsChecked = true, BindingParameter = "Unit" , OrderIndex = 8 },
                    new PreviewColumn() { HeaderName = "Unit Rate", IsChecked = true, BindingParameter = "UnitRate" , OrderIndex = 9 },
                    new PreviewColumn() { HeaderName = "Total Amount", IsChecked = true, BindingParameter = "TotalAmount" , OrderIndex = 10 },
                    new PreviewColumn() { HeaderName = "Specification", IsChecked = true, BindingParameter = "Specification" , OrderIndex = 11 },
                    new PreviewColumn() { HeaderName = "Length", IsChecked = true, BindingParameter = "Length" , OrderIndex = 12 },
                    new PreviewColumn() { HeaderName = "Area", IsChecked = true, BindingParameter = "Area" , OrderIndex = 13 },
                    new PreviewColumn() { HeaderName = "Volume", IsChecked = true, BindingParameter = "Volume" , OrderIndex = 14},
                    new PreviewColumn() { HeaderName = "Count", IsChecked = true, BindingParameter = "Count" , OrderIndex = 15 }
                };

            foreach (PreviewColumn pc in ExportViewModel.ElementsPreviewColumns)
            {
                pc.IsChecked = true;
                var defaultColumn = previewColumns.Find(x => x.HeaderName == pc.HeaderName);
                pc.OrderIndex = defaultColumn.OrderIndex;
            }


            foreach (PreviewColumn pc in ExportViewModel.ElementsPreviewColumns)
            {
                DataGridColumn column = PreviewDataGrid.Columns.Where(x => x.Header.ToString() == pc.HeaderName).FirstOrDefault();
                if (!pc.IsChecked)
                {
                    column.Visibility = Visibility.Collapsed;
                }
                else
                {
                    column.Visibility = Visibility.Visible;
                }
                column.DisplayIndex = pc.OrderIndex;
            }
            ExportViewModel.SaveConfigSet();
        }

        private void PreviewColumnUnchecked(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
            if (checkBox != null)
            {
                PreviewColumn previewColumn = checkBox.DataContext as PreviewColumn;
                if (checkBox.Name == "ElementColumns")
                {
                    DataGridColumn column = PreviewDataGrid.Columns.Where(x => x.Header.ToString() == previewColumn.HeaderName).FirstOrDefault();
                    column.Visibility = Visibility.Collapsed;
                }
                else
                {
                    DataGridColumn column = MaterialsPreviewDataGrid.Columns.Where(x => x.Header.ToString() == previewColumn.HeaderName).FirstOrDefault();
                    column.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void PreviewColumnChecked(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
            if (checkBox != null)
            {
                PreviewColumn previewColumn = checkBox.DataContext as PreviewColumn;
                if (checkBox.Name == "ElementColumns")
                {
                    DataGridColumn column = PreviewDataGrid.Columns.Where(x => x.Header.ToString() == previewColumn.HeaderName).FirstOrDefault();
                    column.Visibility = Visibility.Visible;
                }
                else
                {
                    DataGridColumn column = MaterialsPreviewDataGrid.Columns.Where(x => x.Header.ToString() == previewColumn.HeaderName).FirstOrDefault();
                    column.Visibility = Visibility.Visible;
                }

            }
        }


        bool isAutoReorder = false;
        private void PreviewDataGrid_ColumnReordered(object sender, DataGridColumnEventArgs e)
        {
            if (isAutoReorder) return;
            DataGrid dataGrid = sender as DataGrid;

            foreach (var column in dataGrid.Columns)
            {
                if(dataGrid.Name == "PreviewDataGrid")
                {
                    PreviewColumn previewColumn = ExportViewModel.ElementsPreviewColumns.Where(x => x.HeaderName == column.Header.ToString()).FirstOrDefault();
                    if (previewColumn != null) previewColumn.OrderIndex = column.DisplayIndex;
                }
                else
                {
                    PreviewColumn previewColumn = ExportViewModel.MaterialsPreviewColumns.Where(x => x.HeaderName == column.Header.ToString()).FirstOrDefault();
                    if (previewColumn != null) previewColumn.OrderIndex = column.DisplayIndex;
                }

                
            }
        }

        private void BtnClick_SaveNewConfig(object sender, RoutedEventArgs e)
        {
            ExportViewModel.SaveConfigSet();
        }
    }
}
