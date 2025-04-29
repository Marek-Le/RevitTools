using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Forms = System.Windows.Forms;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class TagQualityDataModel
    {
        public List<TagElementModel> SelectedTagElements { get; set; }
        public List<TagElementModel> TagElements { get; set; }
        public ExportManager ExportManager { get; set; }
        public List<ElementCategory> Categories { get; set; }
        public BackgroundWorker Worker { get; set; }
        public RevitTransform RevitTransform { get; set; }

        string urlGlobal = @"https://gigamap-dev.tesla.com/backend/db/revit_element/add";

        #region Revit Specific Properties and Fields
        public ExternalEvent TheEvent { get; set; }
        public List<Element> SelectedElements { get; set; }
        public TagQualityAction TagQualityAction { get; set; }
        public ManualResetEvent SignalEvent = new ManualResetEvent(false);
        private Document _document;

        #endregion

        private TagQualityDataModel()
        {
            
        }

        public static TagQualityDataModel Initialize(ExternalEvent externalEvent)
        {
            TagQualityDataModel result = new TagQualityDataModel();
            result.TheEvent = externalEvent;
            result.ExportManager = new ExportManager();
            result.TagElements = new List<TagElementModel>();
            result.GetCategories();
            result.GetTransform(); 
            return result;
        }

        //same for navis and revit
        public void SaveAsCsv()
        {
            List<CustomParameter> reportParameters = new List<CustomParameter>()
            {
                new CustomParameter() { Category = "Element", Name = "Workset" },
                new CustomParameter() { Category = "Item", Name = "Source File" },
                new CustomParameter() { Category = "Element", Name = "Family" },
                new CustomParameter() { Category = "Element", Name = "Type" },
                new CustomParameter() { Category = "Element", Name = "Id" },
                new CustomParameter() { Category = "Element", Name = "Category" },
                new CustomParameter() { Category = "Element", Name = "Mark" },
                new CustomParameter() { Category = "Custom", Name = "ARCU-Tag-Number" },
                new CustomParameter() { Category = "Custom", Name = "Equipment Number" },
                new CustomParameter() { Category = "Custom", Name = "PGMM - Designation" },
                new CustomParameter() { Category = "Compare", Name = "Status" },
                new CustomParameter() { Category = "GB", Name = "Location" },
                new CustomParameter() { Category = "GB", Name = "Equipment Type" },
                new CustomParameter() { Category = "GB", Name = "System Type" },
                new CustomParameter() { Category = "GB", Name = "Equipment Code" },
                new CustomParameter() { Category = "GB", Name = "TagCheckStatus" },
                new CustomParameter() { Category = "Grid", Name = "Intersection" },
                new CustomParameter() { Category = "Element", Name = "UniqueId" }
            };
            string header = "";
            Forms.SaveFileDialog saveFileDialog = new Forms.SaveFileDialog();
            saveFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.AddExtension = true;
            if (saveFileDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write), Encoding.UTF8))
                {
                    foreach (CustomParameter cp in reportParameters)
                    {
                        header += "," + cp.Category + " " + cp.Name;
                    }
                    sw.WriteLine(header);

                    foreach (TagElementModel model in TagElements)
                    {
                        string line = "";
                        line += "," + model.Workset.Replace(",", " ");
                        line += "," + model.SourceFile.Replace(",", " ");
                        line += "," + model.FamilyName.Replace(",", " ");
                        line += "," + model.FamilyType.Replace(",", " ");
                        line += "," + model.RevitId.Replace(",", " ");
                        line += "," + model.Category.Replace(",", " ");
                        line += "," + Regex.Replace(model.ElementMark.Replace(",", " "), @"\t|\n|\r", "");
                        line += "," + Regex.Replace(model.ArcuTag.Replace(",", " "), @"\t|\n|\r", "");
                        line += "," + Regex.Replace(model.EquipmentNumber.Replace(",", " "), @"\t|\n|\r", "");
                        line += "," + Regex.Replace(model.PgmmDesignation.Replace(",", " "), @"\t|\n|\r", "");
                        line += "," + Enum.GetName(typeof(CompareStatus), model.CompareStatus);
                        line += "," + model.Location?.Replace(",", " ");
                        line += "," + model.EquipmentCategory?.Replace(",", " ");
                        line += "," + model.SystemName?.Replace(",", " ");
                        line += "," + model.EquipmentCode?.Replace(",", " ");
                        line += "," + Enum.GetName(typeof(TagCheckStatus), model.TagCheckStatus);
                        line += "," + model.GridIntersection.Replace(",", " ");
                        line += "," + model.UniqueId.Replace(",", " ");
                        sw.WriteLine(line);
                    }
                    sw.Flush();
                    sw.Close();
                }
            }
        }

        public void GetTransform()
        {
            RevitTransform = new RevitTransform();
            RevitTransform.CreateCustomTransform(ConvertMmToFeet(new Position(0, 0, 0)), new Position(1, 0, 0), new Position(0, 1, 0), new Position(0, 0, 1));
        }

        public void UpdateSelectedTags(Document doc)
        {
            using (Transaction tx = new Transaction(doc, "Update Marks - Auto"))
            {
                tx.Start();
                foreach (TagElementModel model in SelectedTagElements.Where(e => e.EditStatus == "Modified/Not updated").ToList())
                {
                    Parameter mark = model.Element.get_Parameter(BuiltInParameter.DOOR_NUMBER);
                    mark.Set(model.ElementMark);
                    Parameter arcu = model.Element.LookupParameter("ARCU-Tag-Number");
                    if (arcu != null)
                    {
                        if (model.ArcuTag != "not found") arcu.Set(model.ArcuTag);
                    }
                    model.EditStatus = "Updated";
                }
                tx.Commit();
            }
            SignalEvent.Set();
        }

        public void UpdateAllTags(Document doc)
        {
            using(Transaction tx = new Transaction(doc, "Update Marks - Auto"))
            {
                tx.Start();
                foreach (TagElementModel model in TagElements.Where(e => e.EditStatus == "Modified/Not updated").ToList())
                {
                    Parameter mark = model.Element.get_Parameter(BuiltInParameter.DOOR_NUMBER);
                    mark.Set(model.ElementMark);
                    Parameter arcu = model.Element.LookupParameter("ARCU-Tag-Number");
                    if(arcu != null)
                    {
                        if (model.ArcuTag != "not found") arcu.Set(model.ArcuTag);
                    }
                    model.EditStatus = "Updated";
                }
                tx.Commit();
            }
            SignalEvent.Set();
        }

        private bool DesignOptionNullOrPrimary(Element element)
        {
            bool result = false;
            if(element.DesignOption == null)
            {
                result = true;
            }
            else
            {
                result = element.DesignOption.IsPrimary;
            }
            return result;
        }

        private bool IsSuperComponent(FamilyInstance fi)
        {
            if(fi.SuperComponent == null)
            {
                return true;
            }
            else
            {
                return false;  
            }
        }

        public void FindElementsByCategory(Document doc, bool activeViewOnly, bool excludeSubComponents)
        {            
            _document = doc;
            TagElements.Clear();
            List<int> selectedCategories = Categories.Where(e => e.IsChecked).Select(e => (int)e.BuiltInCategory).ToList();
            List<Element> itemCollection = null;
            if (activeViewOnly)
            {
                itemCollection = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(FamilyInstance)).Where(e => e.Category != null && selectedCategories.Contains(e.Category.Id.IntegerValue) && DesignOptionNullOrPrimary(e)).ToList();
            }
            else
            {
                itemCollection = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Where(e => e.Category != null && selectedCategories.Contains(e.Category.Id.IntegerValue) && DesignOptionNullOrPrimary(e)).ToList();
            }

            if (excludeSubComponents) itemCollection = itemCollection.Where(e => IsSuperComponent(e as FamilyInstance)).ToList();

            int progress = 0;
            int totalItemsFound = 0;
            int counter = 0;

            List<Workset> worksets = new FilteredWorksetCollector(doc).ToList();

            foreach (Element element in itemCollection)
            {
                counter++;

                TagElementModel tagElementModel = TagElementModel.Initialize(element, RevitTransform, doc.Title);
                tagElementModel.GetWorksetName(worksets);
                if (TagMeetsCriteria(tagElementModel.ElementMark))
                {
                    string[] tagCodes = tagElementModel.ElementMark.Replace("BER-", "").Split('-');
                    tagElementModel.ReadTagCodes(tagCodes, ExportManager);
                    tagElementModel.CheckTag();
                }
                tagElementModel.CheckTag();
                TagElements.Add(tagElementModel);
                totalItemsFound++;
                    
                if (counter == 100)
                {
                    progress += 100;
                    counter = 0;
                    double percentage = ((double)progress / (double)itemCollection.Count) * 100;
                    Worker.ReportProgress((int)percentage, string.Format("{0} % (added: {1})", (int)percentage, totalItemsFound));

                }
            }

            FindNearestGrids(TagElements, GridSearchOptions.HostModel, doc);

            List<string> allTags = TagElements.Where(e => e.TagCheckStatus == TagCheckStatus.Tag_OK).Select(e => e.ElementMark.Replace("BER-", "")).ToList();
            foreach (TagElementModel model in TagElements) model.CheckDuplicates(allTags);

            Worker.ReportProgress(100, string.Format("{0} % (added: {1})", 100, totalItemsFound));
            SignalEvent.Set();
        }

        public void SelectAndZoom(UIDocument uidoc)
        {
            uidoc.Selection.SetElementIds(SelectedElements.Select(e => e.Id).ToList());
            uidoc.ShowElements(SelectedElements.Select(e => e.Id).ToList());
        }

        public void SendToGigabase(List<TagElementModel> selectedElements)
        {
            if (selectedElements == null) return;
            string modelGuid = "guid not found";
            if(_document != null && _document.IsModelInCloud) modelGuid = _document.GetCloudModelPath().GetModelGUID().ToString();
            List<ExportedEquipmentData> sendItems = new List<ExportedEquipmentData>();
            foreach (TagElementModel tagModel in selectedElements)
            {
                ExportedEquipmentData data = new ExportedEquipmentData()
                {
                    revit_element_id = tagModel.RevitId,
                    revit_element_guid = tagModel.UniqueId,
                    equipment_tag = tagModel.ElementMark,
                    bounding_box = tagModel.BBoxString,
                    element_data = GetInstanceDataString(tagModel),
                    model_file_name = tagModel.SourceFile,
                    model_file_guid = modelGuid,
                    nearest_grid_intersection = tagModel.GridIntersection
                };
                sendItems.Add(data);
            }
            string jsonString = JsonConvert.SerializeObject(sendItems);
            StringContent httpContent = new StringContent(jsonString, Encoding.UTF8, "application/json");
            HttpClientPost(urlGlobal, httpContent, selectedElements.Count);
        }

        private void HttpClientPost(string url, StringContent stringContent, int count)
        {
            HttpClient client = new HttpClient();
            HttpResponseMessage response = client.PostAsync(url, stringContent).Result;
            string responseBody = response.Content.ReadAsStringAsync().Result;
            if(response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                System.Windows.MessageBox.Show(String.Format("Success sending {0} elements to GigaBase!", count));
            }
            else
            {
                System.Windows.MessageBox.Show("Error while sending objects to GigaBase!");
            }
            System.Windows.MessageBox.Show(responseBody);
        }

        private string GetInstanceDataString(TagElementModel element)
        {
            Element revitElement = element.Element;
            string systemAbbreviation = revitElement.LookupParameter("System Abbreviation")?.AsString();
            string systemClassification = revitElement.LookupParameter("System Classification")?.AsString();
            string systemName = revitElement.LookupParameter("System Name")?.AsString();
            string systemType = revitElement.LookupParameter("System Type")?.AsString();
            string systemTypeId = revitElement.LookupParameter("System Type")?.AsElementId().IntegerValue.ToString();


            return string.Format(@"[{{ ""family_name"": ""{0}"", ""type_name"": ""{1}"", ""category_name"": ""{2}"",
                                        ""sys_abbreviation"": ""{3}"", ""sys_classification"": ""{4}"", ""sys_name"": ""{5}"", ""sys_type"": ""{6}"", ""sys_type_id"": ""{7}"" }}]",
                                         element.FamilyName.Replace('"', 'i'), element.FamilyType.Replace('"', 'i'), element.Category,
                                            systemAbbreviation, systemClassification, systemName, systemType, systemTypeId);
        }

        private bool TagMeetsCriteria(string tag)
        {
            bool result = false;
            string[] tagCodes = tag.Replace("BER-", "").Split('-');
            if (tagCodes.Length == 4)
            {
                string loc = tagCodes[0];
                string eqType = tagCodes[1];
                string sysType = tagCodes[2];
                string id = tagCodes[3];
                if (loc.Length > 1 && loc.Length < 5 && eqType.Length > 1 && eqType.Length < 5 && sysType.Length == 3 && id.Length > 2 && id.Length < 5)
                {
                    result = true;
                }
            }
            return result;
        }

        private void FindNearestGrids(List<TagElementModel> tagElementModels, GridSearchOptions option, Document doc)
        {
            string logData = "";
            #region Find nearest grids in linked model			
            if (option != GridSearchOptions.HostModel)
            {
                List<RevitLinkInstance> gridLinkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Select(e => e as RevitLinkInstance).Where(e => e.Name.Contains("GRID")).ToList();
                if (gridLinkInstances != null && gridLinkInstances.Count == 1)
                {
                    RevitLinkInstance gridLink = gridLinkInstances.FirstOrDefault();
                    Transform gridLinktransform = gridLink.GetTransform();
                    Document gridLinkDoc = gridLink.GetLinkDocument();
                    if (gridLinkDoc != null)
                    {
                        List<Grid> linkGrids = new FilteredElementCollector(gridLinkDoc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().Select(e => e as Grid).ToList();

                        List<Grid> horizontalGrids = new List<Grid>();
                        List<Grid> verticalGrids = new List<Grid>();

                        foreach (Grid grid in linkGrids)
                        {
                            double parsingResult = 0;
                            if (double.TryParse(grid.Name, out parsingResult)) continue;
                            if (IsHorizontal(grid) == true)
                            {
                                horizontalGrids.Add(grid);
                            }
                            if (IsHorizontal(grid) == false)
                            {
                                verticalGrids.Add(grid);
                            }
                        }

                        foreach (TagElementModel model in tagElementModels)
                        {
                            var horderedList = horizontalGrids.OrderBy(e => FindDistance(e, model.Element as FamilyInstance, gridLinktransform)).ToList();
                            var vorderedList = verticalGrids.OrderBy(e => FindDistance(e, model.Element as FamilyInstance, gridLinktransform)).ToList();
                            if (horderedList.Count == 0 || vorderedList.Count == 0)
                            {
                                logData += string.Format("Id: {0} - nearest grid v or h not found in linked model", model.Element.Id.IntegerValue) + Environment.NewLine;
                                continue;
                            }
                            KeyValuePair<Grid, double> hGridDistancePairMin = new KeyValuePair<Grid, double>(horderedList.First(), FindDistance(horderedList.First(), model.Element as FamilyInstance, gridLinktransform));
                            KeyValuePair<Grid, double> vGridDistancePairMin = new KeyValuePair<Grid, double>(vorderedList.First(), FindDistance(vorderedList.First(), model.Element as FamilyInstance, gridLinktransform));
                            model.GetNearestGrids(hGridDistancePairMin, vGridDistancePairMin);
                            model.SourceFile = doc.Title;
                        }
                    }
                    else
                    {
                        logData += "Linked document is null. Try reloading link" + Environment.NewLine;
                    }

                }
                else
                {
                    logData += "Found more than 1 linked grid instance" + Environment.NewLine;
                }
            }
            #endregion
            #region Find nearest grids in host model
            if (option != GridSearchOptions.LinkedModel)
            {
                List<Grid> hostGrids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().Select(e => e as Grid).ToList();
                List<Grid> horizontalGrids = new List<Grid>();
                List<Grid> verticalGrids = new List<Grid>();

                foreach (Grid grid in hostGrids)
                {
                    double parsingResult = 0;
                    if (double.TryParse(grid.Name, out parsingResult)) continue;
                    if (IsHorizontal(grid) == true)
                    {
                        horizontalGrids.Add(grid);
                    }
                    if (IsHorizontal(grid) == false)
                    {
                        verticalGrids.Add(grid);
                    }
                }

                foreach (TagElementModel model in tagElementModels)
                {
                    var horderedList = horizontalGrids.OrderBy(e => FindDistance(e, model.Element as FamilyInstance)).ToList();
                    var vorderedList = verticalGrids.OrderBy(e => FindDistance(e, model.Element as FamilyInstance)).ToList();
                    if (horderedList.Count == 0 || vorderedList.Count == 0) continue;
                    
                    KeyValuePair<Grid, double> hGridDistancePairMin = new KeyValuePair<Grid, double>(horderedList.First(), FindDistance(horderedList.First(), model.Element as FamilyInstance));
                    KeyValuePair<Grid, double> vGridDistancePairMin = new KeyValuePair<Grid, double>(vorderedList.First(), FindDistance(vorderedList.First(), model.Element as FamilyInstance));
                    model.GetNearestGrids(hGridDistancePairMin, vGridDistancePairMin);
                }
            }
            #endregion
        }

        private bool? IsHorizontal(Grid grid)
        {
            Line gridLine = grid.Curve as Line;
            if (gridLine == null)
            {
                TaskDialog.Show("Error", "Grid is not a line!" + Environment.NewLine + grid.Id.IntegerValue.ToString());
            }
            XYZ direction = gridLine.Direction;
            XYZ absDirection = new XYZ(Math.Abs(direction.X), Math.Abs(direction.Y), Math.Abs(direction.Z));
            if (absDirection.IsAlmostEqualTo(new XYZ(1, 0, 0)))
            {
                return true;
            }
            if (absDirection.IsAlmostEqualTo(new XYZ(0, 1, 0)))
            {
                return false;
            }
            else
            {
                return null;
            }
        }

        private double FindDistance(Grid grid, FamilyInstance fi)
        {
            BoundingBoxXYZ box = fi.get_BoundingBox(null);
            XYZ middle = (box.Min + box.Max) / 2;
            Line gridLine = grid.Curve as Line;
            XYZ direction = gridLine.Direction;
            XYZ lineOrigin = gridLine.Origin;
            XYZ start = gridLine.Tessellate()[0];
            XYZ end = gridLine.Tessellate()[1];
            Line gridBound = Line.CreateBound(start, end);
            return gridBound.Distance(middle);
        }

        private double FindDistance(Grid grid, FamilyInstance fi, Transform transform)
        {
            BoundingBoxXYZ box = fi.get_BoundingBox(null);
            XYZ middle = (box.Min + box.Max) / 2;
            Line gridLine = grid.Curve as Line;
            XYZ direction = gridLine.Direction;
            XYZ lineOrigin = transform.OfPoint(gridLine.Origin);
            XYZ lineDirection = transform.OfVector(direction);
            XYZ start = transform.OfPoint(gridLine.Tessellate()[0]);
            XYZ end = transform.OfPoint(gridLine.Tessellate()[1]);
            Line transformedGridBound = Line.CreateBound(start, end);
            return transformedGridBound.Distance(middle);
        }

        private void GetCategories()
        {
            Categories = new List<ElementCategory>();
            Categories.Add(new ElementCategory() { ElementCategoryName = "Air Terminals", IsChecked = true, BuiltInCategory = BuiltInCategory.OST_DuctTerminal });
            Categories.Add(new ElementCategory() { ElementCategoryName = "Duct Accessories", IsChecked = true, BuiltInCategory = BuiltInCategory.OST_DuctAccessory });
            Categories.Add(new ElementCategory() { ElementCategoryName = "Mechanical Equipment", IsChecked = true, BuiltInCategory = BuiltInCategory.OST_MechanicalEquipment });
            Categories.Add(new ElementCategory() { ElementCategoryName = "Pipe Accessories", IsChecked = true, BuiltInCategory = BuiltInCategory.OST_PipeAccessory });
        }

        private Position ConvertMmToFeet(Position position)
        {
            Position result = new Position();
            result.X = Math.Round(position.X / 304.8, 4);
            result.Y = Math.Round(position.Y / 304.8, 4);
            result.Z = Math.Round(position.Z / 304.8, 4);
            return result;
        }

        enum GridSearchOptions
        {
            HostModel,
            LinkedModel,
            Combined
        }
    }
}
