using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows.Input;
using Autodesk.Revit.DB;
using TeslaRevitTools.EquipmentTagQuality.GigabaseModels;


namespace TeslaRevitTools.EquipmentTagQuality
{
    public class TagElementModel : INotifyPropertyChanged
    {
        public Dictionary<string, string> ParameterValues { get; set; } //<parameterName, parameterValue>
        public string SourceFile { get; set; }
        public string Workset { get; set; }
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string FamilyType { get; set; }

        private string _elementMark;
        public string ElementMark
        {
            get => _elementMark;
            set
            {
                if (_elementMark != value)
                {
                    _elementMark = value;
                    OnPropertyChanged(nameof(ElementMark));
                    EditStatus = "Modified/Not updated";
                }
            }
        }

        private string _arcuTag;
        public string ArcuTag
        {
            get => _arcuTag;
            set
            {
                if (_arcuTag != value)
                {
                    _arcuTag = value;
                    OnPropertyChanged(nameof(ArcuTag));
                    EditStatus = "Modified/Not updated";
                }
            }
        }
        public string EquipmentNumber { get; set; }
        public string PgmmDesignation { get; set; }

        private CompareStatus _compareStatus;
        public CompareStatus CompareStatus
        {
            get => _compareStatus;
            set
            {
                if (_compareStatus != value)
                {
                    _compareStatus = value;
                    OnPropertyChanged(nameof(CompareStatus));
                }
            }
        }

        private string _location;
        public string Location
        {
            get => _location;
            set
            {
                if (_location != value)
                {
                    _location = value;
                    OnPropertyChanged(nameof(Location));
                }
            }
        }

        private string _equipmentCategory;
        public string EquipmentCategory
        {
            get => _equipmentCategory;
            set
            {
                if (_equipmentCategory != value)
                {
                    _equipmentCategory = value;
                    OnPropertyChanged(nameof(EquipmentCategory));
                }
            }
        }

        private string _systemName;
        public string SystemName
        {
            get => _systemName;
            set
            {
                if (_systemName != value)
                {
                    _systemName = value;
                    OnPropertyChanged(nameof(SystemName));
                }
            }
        }

        private string _equipmentCode;
        public string EquipmentCode
        {
            get => _equipmentCode;
            set
            {
                if (_equipmentCode != value)
                {
                    _equipmentCode = value;
                    OnPropertyChanged(nameof(EquipmentCode));
                }
            }
        }

        private TagCheckStatus _tagCheckStatus;
        public TagCheckStatus TagCheckStatus
        {
            get => _tagCheckStatus;
            set
            {
                if (_tagCheckStatus != value)
                {
                    _tagCheckStatus = value;
                    OnPropertyChanged(nameof(TagCheckStatus));
                }
            }
        }
        public string UniqueId { get; set; }
        public string RevitId { get; set; }
        public string BBoxString { get; set; }
        public string GridIntersection { get; set; }

        private string _editStatus;
        public string EditStatus
        {
            get => _editStatus;
            set
            {
                if (_editStatus != value)
                {
                    _editStatus = value;
                    OnPropertyChanged(nameof(EditStatus));
                }
            }
        }

        public Element Element { get; set; }

        private string _subcomponent = "";

        //public delegate string CheckTagPartialCode(string partialCode);

        private TagElementModel()
        {

        }

        public string GetHash()
        {
            return SourceFile + Category + FamilyName + FamilyType + ElementMark + ArcuTag + EquipmentNumber + PgmmDesignation + Enum.GetName(typeof(CompareStatus), CompareStatus) + Location + EquipmentCategory + SystemName + EquipmentCode + Enum.GetName(typeof(TagCheckStatus), TagCheckStatus) + UniqueId + RevitId + _subcomponent;
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public static TagElementModel Initialize(Element element, RevitTransform revitTransform, string docName)
        {
            TagElementModel result = new TagElementModel();
            result.ParameterValues = new Dictionary<string, string>();
            result.Element = element;
            result.GetProperties(docName);
            result.GetCompareStatus();
            result.GetBoundingBox(revitTransform);
            result.GetGridIntersection();
            result.EditStatus = "";
            if ((result.Element as FamilyInstance).SuperComponent != null) result._subcomponent = "SUBCOMPONENT";
            return result;
        }

        public void CheckDuplicates(List<string> allTags)
        {
            if (allTags.Where(e => e == ElementMark.Replace("BER-", "")).ToList().Count > 1) 
            { 
                TagCheckStatus = TagCheckStatus.Tag_Duplicated; 
            }
        }

        public void GetCustomParameter(string lookupParameterName)
        {
            Parameter p = Element.LookupParameter(lookupParameterName);
            if(p != null && p.AsString() != null)
            { 
                ParameterValues.Add(lookupParameterName, p.AsString());
            }
            else
            {
                ParameterValues.Add(lookupParameterName, "par. not found");
            }
        }

        public void CopyFromCustomParameter(string parameterName)
        {
            if (string.IsNullOrEmpty(parameterName)) return;
            if(ParameterValues.ContainsKey(parameterName)) 
            {
                ElementMark = ParameterValues[parameterName];
            }
        }

        public void CheckTag()
        {
            string tag = Location + EquipmentCategory + SystemName + EquipmentCode;
            TagCheckStatus = TagCheckStatus.Tag_OK;
            if (string.IsNullOrEmpty(tag)) 
            {
                TagCheckStatus = TagCheckStatus.Tag_Empty;
            }
            else if (tag.Contains("UNKNOWN") || tag.Contains("INCORRECT NUMBER"))
            {
                TagCheckStatus = TagCheckStatus.Tag_Invalid;
            }
        }
    
        public void ReadTagCodes(string[] tagCodes, ExportManager exportManager)
        {
            Location = CheckLocation(tagCodes[0], exportManager);
            EquipmentCategory = CheckEquipmentType(tagCodes[1], exportManager);
            SystemName = CheckSystemType(tagCodes[2], exportManager);
            EquipmentCode = CheckTagCode(tagCodes[3]);
        }

        public void GetCompareStatus()
        {
            bool markOK = TagMeetsCriteria(ElementMark);
            bool arcuOK = TagMeetsCriteria(ArcuTag);

            if (markOK && arcuOK)
            {
                if (ElementMark == ArcuTag)
                {
                    CompareStatus = CompareStatus.Status2;
                }
                else
                {
                    CompareStatus = CompareStatus.Status3;
                }
            }
            if (!markOK && !arcuOK)
            {
                CompareStatus = CompareStatus.Status5;
            }
            if (markOK && !arcuOK)
            {
                CompareStatus = CompareStatus.Status1;
            }
            if (!markOK && arcuOK)
            {
                CompareStatus = CompareStatus.Status4;
            }
        }

        private string CheckLocation(string locationCode, ExportManager exportManager)
        {
            string result = "UNKNOWN";
            EquipmentLocation location = exportManager.locations.Where(e => e.code != null && locationCode.Contains(e.code)).FirstOrDefault();
            if (location != null) result = location.name;
            return result;
        }

        private string CheckEquipmentType(string eqCode, ExportManager exportManager)
        {
            string result = "UNKNOWN";
            EquipmentType eqType = exportManager.eqTypes.Where(e => e.abbreviation != null && e.abbreviation == eqCode).FirstOrDefault();
            if (eqType != null) result = eqType.name;
            return result;
        }

        private string CheckSystemType(string sysCode, ExportManager exportManager)
        {
            string result = "UNKNOWN";
            SystemType sysType = exportManager.sysTypes.Where(e => e.system_code != null && e.system_code == sysCode).FirstOrDefault();
            if (sysType != null) result = sysType.system_name;
            return result;
        }

        private string CheckTagCode(string tagCode)
        {
            string result = "INCORRECT NUMBER";
            if (tagCode.Length == 3)
            {
                if (Int32.TryParse(tagCode, out int integer))
                { 
                    result = "OK"; 
                }
            }
            return result;
        }

        private void GetBoundingBox(RevitTransform revitTransform)
        {

            BoundingBoxXYZ boundingBox = Element.get_BoundingBox(null);

            Position p1 = new Position(boundingBox.Min.X, boundingBox.Min.Y, boundingBox.Min.Z);
            Position p2 = new Position(boundingBox.Max.X, boundingBox.Max.Y, boundingBox.Max.Z);

            Position p1_transformed = revitTransform.TransformAbsolutePosition(p1);
            Position p2_transformed = revitTransform.TransformAbsolutePosition(p2);

            Position boxMin = new Position();
            Position boxMax = new Position();

            if(p1_transformed.X < p2_transformed.X)
            {
                boxMin.X = p1_transformed.X;
                boxMax.X = p2_transformed.X;

            }
            else
            {
                boxMin.X = p2_transformed.X;
                boxMax.X = p1_transformed.X;
            }

            if (p1_transformed.Y < p2_transformed.Y)
            {
                boxMin.Y = p1_transformed.Y;
                boxMax.Y = p2_transformed.Y;
            }
            else
            {
                boxMin.Y = p2_transformed.Y;
                boxMax.Y = p1_transformed.Y;
            }

            if (p1_transformed.Z < p2_transformed.Z)
            {
                boxMin.Z = p1_transformed.Z;
                boxMax.Z = p2_transformed.Z;
            }
            else
            {
                boxMin.Z = p2_transformed.Z;
                boxMax.Z = p1_transformed.Z;
            }

            BBoxString = String.Format("{0} {1} {2} {3} {4} {5}",
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMin.X, UnitTypeId.Millimeters), 0),
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMin.Y, UnitTypeId.Millimeters), 0),
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMin.Z, UnitTypeId.Millimeters), 0),
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMax.X, UnitTypeId.Millimeters), 0),
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMax.Y, UnitTypeId.Millimeters), 0),
                Math.Round(UnitUtils.ConvertFromInternalUnits(boxMax.Z, UnitTypeId.Millimeters), 0));
        }

        public bool TagMeetsCriteria(string tag)
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

        public void GetWorksetName(List<Workset> worksets)
        {
            List<string> names = worksets.Where(e => e.Id == Element.WorksetId).Select(e => e.Name).ToList();

            if (names.Count == 0)
            {
                Workset = "Not assigned";
            }
            else
            {
                Workset = names.First();
            }
        }

        private void GetProperties(string docName)
        {
            SourceFile = docName;
            FamilyName = (Element as FamilyInstance).Symbol.Family.Name;
            FamilyType = (Element as FamilyInstance).Symbol.Name;
            UniqueId = Element.UniqueId;
            Category = Element.Category.Name;
            ElementMark = Element.get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();
            if (string.IsNullOrEmpty(ElementMark)) ElementMark = "not found";
            ArcuTag = GetParameterValue("ARCU-Tag-Number");
            EquipmentNumber = GetParameterValue("Equipment Number");
            PgmmDesignation = GetParameterValue("PGMM - Designation");
            RevitId = Element.Id.IntegerValue.ToString();
        }

        private string GetParameterValue(string name)
        {
            string result = "not found";
            Parameter dataProperty = Element.LookupParameter(name);
            if (dataProperty != null)
            {
                switch (dataProperty.StorageType)
                {
                    case StorageType.None:
                        break;
                    case StorageType.Integer:
                        result = dataProperty.AsInteger().ToString();
                        break;
                    case StorageType.Double:
                        result = dataProperty.AsDouble().ToString("F0");
                        break;
                    case StorageType.String:
                        result = dataProperty.AsString();
                        break;
                    case StorageType.ElementId:
                        result = dataProperty.AsElementId().ToString();
                        break;
                    default:
                        break;
                }
            }
            if(result == null) result = "empty";
            return result;
        }

        private void GetGridIntersection()
        {
            GridIntersection = "---";
            //BoundingBox3D bb3d = ModelItem.BoundingBox();
            //GridSystem oGS = Application.ActiveDocument.Grids.ActiveSystem;
            //if (oGS != null)
            //{

            //    GridIntersection oGridIntersection = oGS.ClosestIntersection(bb3d.Center);
            //    if (oGridIntersection != null)
            //    {
            //        GridIntersection = oGridIntersection.Line1.DisplayName + "/" + oGridIntersection.Line2.DisplayName;
            //    }
            //}
        }

        KeyValuePair<Grid, double> _hGridDistancePairMin;
        KeyValuePair<Grid, double> _vGridDistancePairMin;

        public void GetNearestGrids(KeyValuePair<Grid, double> hGridDistancePairMin, KeyValuePair<Grid, double> vGridDistancePairMin)
        {
            if (_hGridDistancePairMin.Equals(new KeyValuePair<Grid, double>()))
            {
                _hGridDistancePairMin = hGridDistancePairMin;
            }
            else
            {
                if (_hGridDistancePairMin.Value > hGridDistancePairMin.Value)
                {
                    _hGridDistancePairMin = hGridDistancePairMin;
                }
            }

            if (_vGridDistancePairMin.Equals(new KeyValuePair<Grid, double>()))
            {
                _vGridDistancePairMin = vGridDistancePairMin;
            }
            else
            {
                if (_vGridDistancePairMin.Value > vGridDistancePairMin.Value)
                {
                    _vGridDistancePairMin = vGridDistancePairMin;
                }
            }
            GridIntersection = string.Format("{0}/{1}", _hGridDistancePairMin.Key.Name, _vGridDistancePairMin.Key.Name);
        }
    }
}
