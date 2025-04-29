using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using Autodesk.Revit.DB.ExtensibleStorage;
using Forms = System.Windows.Forms;
using SpreadsheetLight;
using Autodesk.Revit.UI;
using System.IO;
using Newtonsoft.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using System.Collections.Specialized;

namespace TeslaRevitTools.ITwoExport
{
    public class ExportViewModel
    {
        public Nullable<bool> HeaderCheckbox { get; set; }
        public ObservableCollection<CategoryModel> RevitCategories { get; set; }
        public ExportedMaterial SelectedMaterial { get; set; }
        public ExportedElement SelectedElement { get; set; }
        public ObservableCollection<PreviewItem> PreviewItems { get; set; }
        public ObservableCollection<PreviewMaterial> PreviewMaterials { get; set; }
        //public ObservableCollection<MaterialItem> MaterialPreviewItems { get; set; }
        public ObservableCollection<ExportedMaterial> ExportedMaterials { get; set; }
        public ObservableCollection<ExportedElement> ExportedElements { get; set; }
        public ObservableCollection<ExcelKeynote> ExcelKeynotes { get; set; }
        public ExternalEvent TheEvent { get; set; }
        public ITwoAction Action { get; set; }
        public bool ActiveViewOnly { get; set; } = false;
        public bool IncludePaintMaterials { get; set; } = true;
        public List<PreviewColumn> ElementsPreviewColumns { get; set; }
        public List<PreviewColumn> MaterialsPreviewColumns { get; set; }
        public Dictionary<string, List<PreviewColumn>> ExportConfigSet { get; set; }

        private ExportViewModel() { }

        public static ExportViewModel Initialize(ExternalEvent theEvent)
        {
            ExportViewModel result = new ExportViewModel();
            result.ExportedElements = new ObservableCollection<ExportedElement>();
            result.ExportedMaterials = new ObservableCollection<ExportedMaterial>();
            result.ExcelKeynotes = new ObservableCollection<ExcelKeynote>();
            result.PreviewItems = new ObservableCollection<PreviewItem>();
            result.PreviewMaterials = new ObservableCollection<PreviewMaterial>();
            result.RevitCategories = new ObservableCollection<CategoryModel>();
            result.TheEvent = theEvent;
            //result.SetDefaultPreviewColumns();
            return result;
        }

        public void SetDefaultPreviewColumns()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "itwoExportSettings.json");
            if (File.Exists(filePath))
            {
                string jsonString = File.ReadAllText(filePath);
                ExportConfigSet = JsonConvert.DeserializeObject<Dictionary<string, List<PreviewColumn>>>(jsonString);
                ElementsPreviewColumns = ExportConfigSet["defaultElements"];
                MaterialsPreviewColumns = ExportConfigSet["defaultMaterials"];
            }
            else
            {
                ElementsPreviewColumns = new List<PreviewColumn>
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
                MaterialsPreviewColumns = new List<PreviewColumn>
                {
                    new PreviewColumn() { HeaderName = "Material Name", IsChecked = true, BindingParameter = "MaterialName", OrderIndex = 0},
                    new PreviewColumn() { HeaderName = "Description", IsChecked = true, BindingParameter = "Description", OrderIndex = 1  },
                    new PreviewColumn() { HeaderName = "KeyNotes", IsChecked = true, BindingParameter = "Keynote", OrderIndex = 2  },
                    new PreviewColumn() { HeaderName = "RN", IsChecked = true, BindingParameter = "Rn", OrderIndex = 3  },
                    new PreviewColumn() { HeaderName = "Short Info", IsChecked = true, BindingParameter = "ShortInfo", OrderIndex = 4  },
                    new PreviewColumn() { HeaderName = "Outl. Spec.", IsChecked = true, BindingParameter = "OutlineSpec", OrderIndex = 5 },
                    new PreviewColumn() { HeaderName = "Quantity", IsChecked = true, BindingParameter = "Quantity" , OrderIndex = 6 },
                    new PreviewColumn() { HeaderName = "UoM", IsChecked = true, BindingParameter = "Unit" , OrderIndex = 7 },
                    new PreviewColumn() { HeaderName = "Unit Rate", IsChecked = true, BindingParameter = "UnitRate" , OrderIndex = 8 },
                    new PreviewColumn() { HeaderName = "Total Amount", IsChecked = true, BindingParameter = "TotalAmount" , OrderIndex = 9 },
                    new PreviewColumn() { HeaderName = "Specification", IsChecked = true, BindingParameter = "Specification" , OrderIndex = 10 },
                    new PreviewColumn() { HeaderName = "Area", IsChecked = true, BindingParameter = "Area" , OrderIndex = 11 },
                    new PreviewColumn() { HeaderName = "Volume", IsChecked = true, BindingParameter = "Volume" , OrderIndex = 12},
                };
                ExportConfigSet = new Dictionary<string, List<PreviewColumn>>();
                ExportConfigSet.Add("defaultElements", ElementsPreviewColumns);
                ExportConfigSet.Add("defaultMaterials", MaterialsPreviewColumns);
                string jsonString = JsonConvert.SerializeObject(ExportConfigSet);
                File.WriteAllText(filePath, jsonString);
            }
        }

        public void SaveConfigSet()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "itwoExportSettings.json");
            string jsonString = JsonConvert.SerializeObject(ExportConfigSet);
            File.WriteAllText(filePath, jsonString);
        }

        public void ResetDataGrids()
        {
            ExportedElements.Clear();
            ExportedMaterials.Clear();
            PreviewItems.Clear();
            PreviewMaterials.Clear();
        }

        public void TryGetOptions()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "iTwoOptions.json");
            if (File.Exists(filePath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(filePath);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string content = lines[i];
                        if(i == 1)
                        {
                            if (content == "False")
                            {
                                ActiveViewOnly = false;
                            }
                            else if (content == "True")
                            {
                                ActiveViewOnly = true;
                            }
                            else
                            {
                                File.Delete(filePath);
                            }
                        }
                        if (i == 2)
                        {
                            if (content == "False")
                            {
                                IncludePaintMaterials = false;
                            }
                            else if (content == "True")
                            {
                                IncludePaintMaterials = true;
                            }
                            else
                            {
                                File.Delete(filePath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.ToString());
                }
            }
        }

        public ObservableCollection<CategoryModel> TryGetCategories()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "iTwoSelectedCategories.json");
            if (File.Exists(filePath))
            {
                try
                {
                    ObservableCollection<CategoryModel> categories = JsonConvert.DeserializeObject<ObservableCollection<CategoryModel>>(File.ReadAllText(filePath));
                    return categories;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.ToString());
                    return null;
                }
            }
            return null;
        }

        public void TryGetExcelKeynotes()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "iTwoExcelData.json");
            if(File.Exists(filePath))
            {
                try
                {
                    ExcelKeynotes = JsonConvert.DeserializeObject<ObservableCollection<ExcelKeynote>>(File.ReadAllText(filePath));
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.ToString());
                }
            }
        }

        public void SelectElementContaining(UIDocument uidoc)
        {
            if (SelectedMaterial != null || SelectedMaterial.Instance.IsValidObject)
            {
                uidoc.Selection.SetElementIds(new List<ElementId>() { SelectedMaterial.Instance.Id });
            }
        }

        public void SelectInModel(UIDocument uidoc)
        {
            if (SelectedElement != null || SelectedElement.Instance.IsValidObject)
            {
                uidoc.Selection.SetElementIds(new List<ElementId>() { SelectedElement.Instance.Id });
            }
        }

        public void SelectMaterial(UIDocument uidoc)
        {
            if (SelectedMaterial != null || SelectedMaterial.Material.IsValidObject)
            {
                uidoc.Selection.SetElementIds(new List<ElementId>() { SelectedMaterial.Material.Id });
            }
        }

        public void CheckCategories(Document doc)
        {
            RevitCategories.Clear();
            List<CategoryModel> categories = new List<CategoryModel>();
            foreach (Category c in doc.Settings.Categories)
            {
                if (c.CategoryType == CategoryType.Model)
                {
                    if (string.IsNullOrEmpty(c.Name)) continue;
                    CategoryModel category = new CategoryModel()
                    {
                        IsSelected = true,
                        Name = c.Name,
                        BuiltInCategory = c.BuiltInCategory
                    };
                    categories.Add(category);
                }
            }
            categories = categories.OrderBy(e => e.Name).ToList();
            var checkList = TryGetCategories();
            foreach (var item in categories)
            {
                if (checkList != null && checkList.Count > 0)
                {
                    List<int> selectedCategories = checkList.Where(e => e.IsSelected).Select(e => (int)e.BuiltInCategory).ToList();
                    if (!selectedCategories.Contains((int)item.BuiltInCategory)) item.IsSelected = false;

                    var check = checkList.Where(e => e.Name == item.Name).FirstOrDefault();
                    if(check != null) item.Density = check.Density;
                }
                RevitCategories.Add(item);
            }
        }

        public void GetExportedMaterials(Document doc)
        {
            ExportedMaterials.Clear();
            List<Element> elements = null;

            if (ActiveViewOnly)
            {
                elements = new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().Where(e => IsOfSelectedCategory(e)).ToList();
            }
            else
            {
                elements = new FilteredElementCollector(doc).WhereElementIsNotElementType().Where(e => IsOfSelectedCategory(e)).ToList();
            }

            foreach (Element e in elements)
            {
                var catModel = RevitCategories.Where(x => (int)x.BuiltInCategory == e.Category.Id.IntegerValue).FirstOrDefault();

                foreach (ElementId id in e.GetMaterialIds(false))
                {
                    //materials in structure
                    Material m = doc.GetElement(id) as Material;
                    ExportedMaterial exportedMaterial = ExportedMaterial.Initialize(e, m, doc, false);
                    if(catModel != null) exportedMaterial.Density = catModel.Density;
                    ExportedMaterials.Add(exportedMaterial);
                }
                if(IncludePaintMaterials)
                {
                    foreach (ElementId id in e.GetMaterialIds(true))
                    {
                        //materials painted on faces
                        Material m = doc.GetElement(id) as Material;
                        ExportedMaterial exportedMaterial = ExportedMaterial.Initialize(e, m, doc, true);
                        if (catModel != null) exportedMaterial.Density = catModel.Density;
                        ExportedMaterials.Add(exportedMaterial);
                    }
                }
            }
        }

        public void GetExportedElements(Document doc)
        {
            ExportedElements.Clear();
            List<Element> elements = null;
            if(ActiveViewOnly)
            {
                elements = new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().Where(e => IsOfSelectedCategory(e)).ToList();
            }
            else
            {
                elements = new FilteredElementCollector(doc).WhereElementIsNotElementType().Where(e => IsOfSelectedCategory(e)).ToList();
            }

            foreach (Element e in elements)
            {
                //TODO: get density and pass to initialize
                var catModel = RevitCategories.Where(x => (int)x.BuiltInCategory == e.Category.Id.IntegerValue).FirstOrDefault();
                ExportedElement exportedElement = ExportedElement.Initialize(e, doc);
                if (exportedElement != null)
                {
                    if(catModel != null) exportedElement.Density = catModel.Density;
                    ExportedElements.Add(exportedElement);
                }
            }
        }

        List<string> _excelKeynotes;

        public void CheckElementKeynotes()
        {
            _excelKeynotes = ExcelKeynotes.Select(e => e.Index0).ToList();
            foreach (ExportedElement exportedElement in ExportedElements) exportedElement.CheckKeynoteExists(_excelKeynotes);
        }

        public void CheckMaterialKeynotes()
        {
            _excelKeynotes = ExcelKeynotes.Select(e => e.Index0).ToList();
            foreach (ExportedMaterial exportedMaterial in ExportedMaterials) exportedMaterial.CheckKeynoteExists(_excelKeynotes);
        }

        public void SaveElementsToExcel()
        {
            Forms.SaveFileDialog saveFileDialog = new Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.AddExtension = true;
            if (saveFileDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                //https://spreadsheetlight.com/sample-code/
                string path = saveFileDialog.FileName;
                SLDocument s1 = new SLDocument();
                string sheetName = s1.GetCurrentWorksheetName();
                s1.RenameWorksheet(sheetName, "List");

                int columnCorrection = -1;

                foreach (var column in ElementsPreviewColumns.OrderBy(e => e.OrderIndex).ToList())
                {
                    if (column.IsChecked)
                    {
                        int columnIndex = column.OrderIndex - columnCorrection;
                        s1.SetCellValue(1, columnIndex, column.HeaderName);
                        for (int i = 0; i < PreviewItems.Count; i++)
                        {
                            PreviewItem previewItem = PreviewItems[i];
                            int rowIndex = i + 2;
                            switch (column.HeaderName)
                            {
                                case "Type Name":
                                    if (!string.IsNullOrEmpty(previewItem.TypeName)) s1.SetCellValue(rowIndex, columnIndex, previewItem.TypeName);
                                    break;
                                case "Description":
                                    if (!string.IsNullOrEmpty(previewItem.Description)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Description);
                                    break;
                                case "KeyNotes":
                                    if (!string.IsNullOrEmpty(previewItem.Keynote)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Keynote);
                                    break;
                                case "DIN":
                                    if (!string.IsNullOrEmpty(previewItem.Din)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Din);
                                    break;
                                case "RN":
                                    if (!string.IsNullOrEmpty(previewItem.Rn)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Rn);
                                    break;
                                case "Short Info":
                                    if (!string.IsNullOrEmpty(previewItem.ShortInfo)) s1.SetCellValue(rowIndex, columnIndex, previewItem.ShortInfo);
                                    break;
                                case "Outl. Spec.":
                                    if (!string.IsNullOrEmpty(previewItem.OutlineSpec)) s1.SetCellValue(rowIndex, columnIndex, previewItem.OutlineSpec);
                                    break;
                                case "Quantity":
                                    if (!string.IsNullOrEmpty(previewItem.Quantity)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Quantity);
                                    break;
                                case "UoM":
                                    if (!string.IsNullOrEmpty(previewItem.Unit)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Unit);
                                    break;
                                case "Unit Rate":
                                    if (!string.IsNullOrEmpty(previewItem.UnitRate)) s1.SetCellValue(rowIndex, columnIndex, previewItem.UnitRate);
                                    break;
                                case "Total Amount":
                                    if (!string.IsNullOrEmpty(previewItem.TotalAmount)) s1.SetCellValue(rowIndex, columnIndex, previewItem.TotalAmount);
                                    break;
                                case "Specification":
                                    if (!string.IsNullOrEmpty(previewItem.Specification)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Specification);
                                    break;
                                case "Length":
                                    if (previewItem.Length > 0) s1.SetCellValue(rowIndex, columnIndex, previewItem.Length);
                                    break;
                                case "Area":
                                    if (!string.IsNullOrEmpty(previewItem.Area)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Area);
                                    break;
                                case "Volume":
                                    if (previewItem.Volume > 0) s1.SetCellValue(rowIndex, columnIndex, previewItem.Volume);
                                    break;
                                case "Count":
                                    if (!string.IsNullOrEmpty(previewItem.Count)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Count);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    else
                    {
                        columnCorrection++;
                    }
                }

                s1.SaveAs(path);
            }
        }

        public void SaveMaterialsToExcel()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.AddExtension = true;
            if (saveFileDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                //https://spreadsheetlight.com/sample-code/
                string path = saveFileDialog.FileName;
                SLDocument s1 = new SLDocument();
                string sheetName = s1.GetCurrentWorksheetName();
                s1.RenameWorksheet(sheetName, "List");

                int columnCorrection = -1;

                foreach (var column in MaterialsPreviewColumns.OrderBy(e => e.OrderIndex).ToList())
                {
                    if (column.IsChecked)
                    {
                        int columnIndex = column.OrderIndex - columnCorrection;
                        s1.SetCellValue(1, columnIndex, column.HeaderName);
                        for (int i = 0; i < PreviewMaterials.Count; i++)
                        {
                            PreviewMaterial previewItem = PreviewMaterials[i];
                            int rowIndex = i + 2;
                            switch (column.HeaderName)
                            {
                                case "Material Name":
                                    if (!string.IsNullOrEmpty(previewItem.MaterialName)) s1.SetCellValue(rowIndex, columnIndex, previewItem.MaterialName);
                                    break;
                                case "Description":
                                    if (!string.IsNullOrEmpty(previewItem.Description)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Description);
                                    break;
                                case "KeyNotes":
                                    if (!string.IsNullOrEmpty(previewItem.Keynote)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Keynote);
                                    break;
                                case "RN":
                                    if (!string.IsNullOrEmpty(previewItem.Rn)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Rn);
                                    break;
                                case "Short Info":
                                    if (!string.IsNullOrEmpty(previewItem.ShortInfo)) s1.SetCellValue(rowIndex, columnIndex, previewItem.ShortInfo);
                                    break;
                                case "Outl. Spec.":
                                    if (!string.IsNullOrEmpty(previewItem.OutlineSpec)) s1.SetCellValue(rowIndex, columnIndex, previewItem.OutlineSpec);
                                    break;
                                case "Quantity":
                                    if (!string.IsNullOrEmpty(previewItem.Quantity)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Quantity);
                                    break;
                                case "UoM":
                                    if (!string.IsNullOrEmpty(previewItem.Unit)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Unit);
                                    break;
                                case "Unit Rate":
                                    if (!string.IsNullOrEmpty(previewItem.UnitRate)) s1.SetCellValue(rowIndex, columnIndex, previewItem.UnitRate);
                                    break;
                                case "Total Amount":
                                    if (!string.IsNullOrEmpty(previewItem.MaterialName)) s1.SetCellValue(rowIndex, columnIndex, previewItem.TotalAmount);
                                    break;
                                case "Specification":
                                    if (!string.IsNullOrEmpty(previewItem.Specification)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Specification);
                                    break;
                                case "Area":
                                    if (!string.IsNullOrEmpty(previewItem.Area)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Area);
                                    break;
                                case "Volume":
                                    if (!string.IsNullOrEmpty(previewItem.Volume)) s1.SetCellValue(rowIndex, columnIndex, previewItem.Volume);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    else
                    {
                        columnCorrection++;
                    }
                }
                s1.SaveAs(path);
            }
        }

        public void SetPreviewItems()
        {
            int rn0 = 0;
            int rn1 = 0;
            int rn2 = 0;
            int rn3 = 0;
            int rn4 = 0;

            PreviewItems.Clear();
            CheckElementKeynotes();

            List<ExportedElement> exportedElements = ExportedElements.Where(e => e.IsKeynoteOk).ToList();
            List<string> mainKeynotes = exportedElements.Select(e => e.Key0).Distinct().OrderBy(e => int.Parse(e)).ToList();
            foreach (string mainKey in mainKeynotes)
            {
                PreviewItem mainItem = new PreviewItem() { Keynote = mainKey };
                ExcelKeynote mainExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == mainKey).FirstOrDefault();
                if (mainExcelKeynote != null)
                {
                    rn0++;
                    rn1 = 0;
                    rn2 = 0;
                    rn3 = 0;
                    rn4 = 0;
                    mainItem.OutlineSpec = mainExcelKeynote.Index1;
                    mainItem.Rn = rn0.ToString("D3");
                }
                else
                {
                    continue;
                }
                PreviewItems.Add(mainItem);
                List<string> secondKeynotes = exportedElements.Where(e => e.Key0 == mainKey).Select(e => e.Key1).Distinct().OrderBy(e => int.Parse(e)).ToList();
                foreach (string secondKey in secondKeynotes)
                {
                    string keynote2 = String.Format("{0}.{1}", mainKey, secondKey);
                    PreviewItem secondItem = new PreviewItem() { Keynote = keynote2 };
                    ExcelKeynote secondExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote2).FirstOrDefault();
                    if (secondExcelKeynote != null) secondItem.OutlineSpec = secondExcelKeynote.Index1;
                    rn1++;
                    rn2 = 0;
                    rn3 = 0;
                    rn4 = 0;
                    if (_excelKeynotes.Contains(keynote2))
                    {
                        string rn1_text = string.Format("{0}.{1}", rn0.ToString("D3"), rn1.ToString("D3"));
                        secondItem.Rn = rn1_text;
                        PreviewItems.Add(secondItem);
                    }

                    List<string> thirdKeynotes = exportedElements.Where(e => e.Key0 == mainKey && e.Key1 == secondKey).Select(e => e.Key2).Distinct().OrderBy(e => int.Parse(e)).ToList();
                    foreach (string thirdKey in thirdKeynotes)
                    {
                        string keynote3 = String.Format("{0}.{1}.{2}", mainKey, secondKey, thirdKey);
                        PreviewItem thirdItem = new PreviewItem() { Keynote =  keynote3};
                        ExcelKeynote thirdExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote3).FirstOrDefault();
                        if (thirdExcelKeynote != null) thirdItem.OutlineSpec = thirdExcelKeynote.Index1;
                        rn2++;
                        rn3 = 0;
                        rn4 = 0;
                        if (_excelKeynotes.Contains(keynote3))
                        {
                            string rn2_text = string.Format("{0}.{1}.{2}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"));
                            thirdItem.Rn = rn2_text;
                            PreviewItems.Add(thirdItem);
                        }

                        List<string> fourthKeynotes = exportedElements.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey).Select(e => e.Key3).Distinct().OrderBy(e => int.Parse(e)).ToList();
                        foreach (string fourthKey in fourthKeynotes)
                        {
                            string keynote4 = String.Format("{0}.{1}.{2}.{3}", mainKey, secondKey, thirdKey, fourthKey);
                            PreviewItem fourthItem = new PreviewItem() { Keynote = keynote4 };
                            ExcelKeynote fourthExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote4).FirstOrDefault();
                            if (fourthExcelKeynote != null) fourthItem.OutlineSpec = fourthExcelKeynote.Index1;
                            rn3++;
                            rn4 = 0;
                            if (_excelKeynotes.Contains(keynote4))
                            {
                                string rn3_text = string.Format("{0}.{1}.{2}.{3}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"), rn3.ToString("D3"));
                                fourthItem.Rn = rn3_text;
                                PreviewItems.Add(fourthItem);
                            }

                            List<ExportedElement> sortedElements = ExportedElements.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey && e.Key3 == fourthKey).OrderBy(e => int.Parse(e.Key4)).ToList();
                            List<string> fifthKeynotes = exportedElements.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey && e.Key3 == fourthKey).Select(e => e.Key4).Distinct().OrderBy(e => int.Parse(e)).ToList();

                            foreach (string fifthKey in fifthKeynotes)
                            {
                                List<ExportedElement> groupedElements = sortedElements.Where(e => e.Key4 == fifthKey).ToList();
                                Parameter par1 = groupedElements.FirstOrDefault().InstanceType.LookupParameter("TSLA_Mass of 1 meter");
                                double density = groupedElements.FirstOrDefault().Density;
                                double massPlength = 0;
                                if (par1 != null) massPlength = UnitUtils.ConvertFromInternalUnits(par1.AsDouble(), UnitTypeId.KilogramsPerMeter);
                                string keynote = String.Format("{0}.{1}.{2}.{3}.{4}", mainKey, secondKey, thirdKey, fourthKey, fifthKey);
                                ExcelKeynote excelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote).FirstOrDefault();

                                var volume = groupedElements.Select(e => e.Volume).Sum();

                                PreviewItem previewItem = new PreviewItem()
                                {
                                    TypeName = groupedElements.FirstOrDefault().InstanceType.Name,
                                    Description = groupedElements.FirstOrDefault().Description,
                                    Keynote = keynote,
                                    Count = groupedElements.Count().ToString(),
                                    Length = groupedElements.Select(e => e.Length).Sum(),
                                    Volume = volume,
                                    Area = groupedElements.Select(e => e.Area).Sum().ToString("F3"),                            
                                };
                                if(excelKeynote != null)
                                {
                                    previewItem.OutlineSpec = excelKeynote.Index1;
                                    previewItem.Unit = excelKeynote.Index3;
                                    previewItem.Specification = excelKeynote.Index4;
                                    switch (previewItem.Unit)
                                    {
                                        case "nos.":
                                            previewItem.Quantity = previewItem.Count;
                                            break;
                                        case "m":
                                            //get length
                                            previewItem.Quantity = previewItem.Length.ToString("F3");
                                            break;
                                        case "m2":
                                            previewItem.Quantity = previewItem.Area;
                                            break;
                                        case "m3":
                                            if(previewItem.Volume != 0) previewItem.Quantity = previewItem.Volume.ToString("F3");
                                            break;
                                        case "kg":
                                            //get mass L*kg/m or V*dens
                                            if(groupedElements.FirstOrDefault().Instance.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Rebar)
                                            {
                                                previewItem.Quantity = (previewItem.Length * massPlength).ToString("F3");
                                            }
                                            else
                                            {
                                                previewItem.Quantity = (previewItem.Volume * density).ToString("F3");
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    Parameter ac = groupedElements.FirstOrDefault().InstanceType.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);
                                    if(ac != null && ac.AsString() != null)
                                    {
                                        previewItem.Din = ac.AsString();
                                    }
                                }
                                rn4++;
                                string rn4_text = string.Format("{0}.{1}.{2}.{3}.{4}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"), rn3.ToString("D3"), rn4.ToString("D3"));
                                previewItem.Rn = rn4_text;
                                PreviewItems.Add(previewItem);
                            }
                        }
                    }
                }
            }
        }

        public void SetPreviewMaterials()
        {
            int rn0 = 0;
            int rn1 = 0;
            int rn2 = 0;
            int rn3 = 0;
            int rn4 = 0;

            PreviewMaterials.Clear();
            CheckMaterialKeynotes();

            List<ExportedMaterial> exportedMaterials = ExportedMaterials.Where(e => e.IsKeynoteOk).ToList();
            List<string> mainKeynotes = exportedMaterials.Select(e => e.Key0).Distinct().OrderBy(e => int.Parse(e)).ToList();
            foreach (string mainKey in mainKeynotes)
            {
                PreviewMaterial mainItem = new PreviewMaterial() { Keynote = mainKey };
                ExcelKeynote mainExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == mainKey).FirstOrDefault();
                if (mainExcelKeynote != null)
                {
                    rn0++;
                    rn1 = 0;
                    rn2 = 0;
                    rn3 = 0;
                    rn4 = 0;
                    mainItem.OutlineSpec = mainExcelKeynote.Index1;
                    mainItem.Rn = rn0.ToString("D3");
                }
                else
                {
                    continue;
                }
                PreviewMaterials.Add(mainItem);
                List<string> secondKeynotes = exportedMaterials.Where(e => e.Key0 == mainKey).Select(e => e.Key1).Distinct().OrderBy(e => int.Parse(e)).ToList();
                foreach (string secondKey in secondKeynotes)
                {
                    string keynote2 = String.Format("{0}.{1}", mainKey, secondKey);
                    PreviewMaterial secondItem = new PreviewMaterial() { Keynote = keynote2 };
                    ExcelKeynote secondExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote2).FirstOrDefault();
                    if (secondExcelKeynote != null) secondItem.OutlineSpec = secondExcelKeynote.Index1;
                    rn1++;
                    rn2 = 0;
                    rn3 = 0;
                    rn4 = 0;
                    if (_excelKeynotes.Contains(keynote2))
                    {
                        string rn1_text = string.Format("{0}.{1}", rn0.ToString("D3"), rn1.ToString("D3"));
                        secondItem.Rn = rn1_text;
                        PreviewMaterials.Add(secondItem);
                    }

                    List<string> thirdKeynotes = exportedMaterials.Where(e => e.Key0 == mainKey && e.Key1 == secondKey).Select(e => e.Key2).Distinct().OrderBy(e => int.Parse(e)).ToList();
                    foreach (string thirdKey in thirdKeynotes)
                    {
                        string keynote3 = String.Format("{0}.{1}.{2}", mainKey, secondKey, thirdKey);
                        PreviewMaterial thirdItem = new PreviewMaterial() { Keynote = keynote3 };
                        ExcelKeynote thirdExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote3).FirstOrDefault();
                        if (thirdExcelKeynote != null) thirdItem.OutlineSpec = thirdExcelKeynote.Index1;
                        rn2++;
                        rn3 = 0;
                        rn4 = 0;
                        if (_excelKeynotes.Contains(keynote3))
                        {
                            string rn2_text = string.Format("{0}.{1}.{2}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"));
                            thirdItem.Rn = rn2_text;
                            PreviewMaterials.Add(thirdItem);
                        }

                        List<string> fourthKeynotes = exportedMaterials.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey).Select(e => e.Key3).Distinct().OrderBy(e => int.Parse(e)).ToList();
                        foreach (string fourthKey in fourthKeynotes)
                        {
                            string keynote4 = String.Format("{0}.{1}.{2}.{3}", mainKey, secondKey, thirdKey, fourthKey);
                            PreviewMaterial fourthItem = new PreviewMaterial() { Keynote = keynote4 };
                            ExcelKeynote fourthExcelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote4).FirstOrDefault();
                            if (fourthExcelKeynote != null) fourthItem.OutlineSpec = fourthExcelKeynote.Index1;
                            rn3++;
                            rn4 = 0;
                            if (_excelKeynotes.Contains(keynote4))
                            {
                                string rn3_text = string.Format("{0}.{1}.{2}.{3}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"), rn3.ToString("D3"));
                                fourthItem.Rn = rn3_text;
                                PreviewMaterials.Add(fourthItem);
                            }

                            List<ExportedMaterial> sortedElements = ExportedMaterials.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey && e.Key3 == fourthKey).OrderBy(e => int.Parse(e.Key4)).ToList();
                            List<string> fifthKeynotes = exportedMaterials.Where(e => e.Key0 == mainKey && e.Key1 == secondKey && e.Key2 == thirdKey && e.Key3 == fourthKey).Select(e => e.Key4).Distinct().OrderBy(e => int.Parse(e)).ToList();

                            foreach (string fifthKey in fifthKeynotes)
                            {
                                List<ExportedMaterial> groupedElements = sortedElements.Where(e => e.Key4 == fifthKey).ToList();
                                string keynote = String.Format("{0}.{1}.{2}.{3}.{4}", mainKey, secondKey, thirdKey, fourthKey, fifthKey);
                                ExcelKeynote excelKeynote = ExcelKeynotes.Where(e => e.Index0 == keynote).FirstOrDefault();
                                double volume = groupedElements.Select(e => e.Volume).Sum();
                                double density = groupedElements.FirstOrDefault().Density;
                                PreviewMaterial previewItem = new PreviewMaterial()
                                {
                                    MaterialName = groupedElements.FirstOrDefault().Material.Name,
                                    Description = groupedElements.FirstOrDefault().Description,
                                    Keynote = keynote,
                                    Volume = volume.ToString("F2"),                                    
                                    Area = groupedElements.Select(e => e.Area).Sum().ToString("F2"),
                                };
                                if (excelKeynote != null)
                                {
                                    previewItem.OutlineSpec = excelKeynote.Index1;
                                    previewItem.Unit = excelKeynote.Index3;
                                    previewItem.Specification = excelKeynote.Index4;
                                    switch (previewItem.Unit)
                                    {
                                        case "m":
                                            break;
                                        case "m2":
                                            previewItem.Quantity = previewItem.Area;
                                            break;
                                        case "m3":
                                            previewItem.Quantity = previewItem.Volume;
                                            break;
                                        case "kg":
                                            previewItem.Quantity = (volume* density).ToString("F2");
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                rn4++;
                                string rn4_text = string.Format("{0}.{1}.{2}.{3}.{4}", rn0.ToString("D3"), rn1.ToString("D3"), rn2.ToString("D3"), rn3.ToString("D3"), rn4.ToString("D3"));
                                previewItem.Rn = rn4_text;
                                PreviewMaterials.Add(previewItem);
                            }
                        }
                    }
                }
            }
        }

        public void GetExcelData(bool clearCurrentData)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select excelfile:";
            openFileDialog.Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|" + "All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog.FileName;
                DataSet dataSet = ExcelMinWrapper.GetExcelDataSet(path);
                if(clearCurrentData) ExcelKeynotes.Clear();

                foreach (DataTable datatable in dataSet.Tables)
                {
                    if (datatable.TableName == "Keynotes")
                    {
                        foreach (DataRow dataRow in datatable.Rows)
                        {
                            ExcelKeynote excelKeynote = new ExcelKeynote();
                            excelKeynote.Index0 = dataRow.ItemArray[0].ToString();
                            excelKeynote.Index1 = dataRow.ItemArray[1].ToString();
                            excelKeynote.Index2 = dataRow.ItemArray[2].ToString();
                            excelKeynote.Index3 = dataRow.ItemArray[3].ToString();
                            excelKeynote.Index4 = dataRow.ItemArray[4].ToString();
                            ExcelKeynotes.Add(excelKeynote);
                        }
                    }
                }
            }
          
        }

        public void SaveExcelDataAsJson()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "iTwoExcelData.json");
            string jsonContent = JsonConvert.SerializeObject(ExcelKeynotes);
            File.WriteAllText(filePath, jsonContent);
            TaskDialog.Show("Info", "Excel data saved!");
        }

        public void SaveSelectedCategories()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            string filePath = Path.Combine(tslaRevitToolsDirectory, "iTwoSelectedCategories.json");
            string jsonContent = JsonConvert.SerializeObject(RevitCategories);
            string optionsPath = Path.Combine(tslaRevitToolsDirectory, "iTwoOptions.json");
            string[] lines = new string[3];
            lines[0] = "Options:";
            lines[1] = ActiveViewOnly.ToString();
            lines[2] = IncludePaintMaterials.ToString();
            File.WriteAllLines(optionsPath, lines);

            File.WriteAllText(filePath, jsonContent);
            TaskDialog.Show("Info", "Configuration saved!");
        }

        private bool IsOfSelectedCategory(Element element)
        {
            List<int> selectedCategories = RevitCategories.Where(e => e.IsSelected).Select(e => (int)e.BuiltInCategory).ToList();
            if(element.Category != null && selectedCategories.Contains(element.Category.Id.IntegerValue))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
