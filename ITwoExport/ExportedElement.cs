using Autodesk.Revit.DB;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.ITwoExport
{
    public class ExportedElement : INotifyPropertyChanged
    {
        public string Hash { get; set; }
        public ElementId InstanceId { get; set; }
        public Element Instance { get; set; }
        public Element InstanceType { get; set; }
        public string Description { get; set; }
        public string CategoryName { get; set; }
        public double Area { get; set; }
        public double Volume { get; set; } //dens 7850 kg/m3
        public double Length { get; set; }

        private string _keynote;
        public string Keynote
        {
            get => _keynote;
            set
            {
                if (_keynote != value)
                {
                    _keynote = value;
                    OnPropertyChanged(nameof(Keynote));
                }
            }
        }
        public string Key0 { get; set; }
        public string Key1 { get; set; }
        public string Key2 { get; set; }
        public string Key3 { get; set; }
        public string Key4 { get; set; }
        public bool IsKeynoteOk { get; set; }
        private KeynoteStatus _keynoteStatus;
        public KeynoteStatus KeynoteStatus
        {
            get => _keynoteStatus;
            set
            {
                if (_keynoteStatus != value)
                {
                    _keynoteStatus = value;
                    OnPropertyChanged(nameof(KeynoteStatus));
                }
            }
        }
        public double Density { get; set; }

        private Document _doc;

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static ExportedElement Initialize(Element element, Document doc)
        {
            ExportedElement result = new ExportedElement();
            result.Instance = element;
            result.InstanceId = element.Id;
            result._doc = doc;
            if(!result.GetInstanceType()) return null;
            if(!result.GetKeynote()) return null;
            result.GetCategoryName();
            result.TrySplitKeynote();
            result.GetDimensions();
            result.GetHash();
            return result;
        }

        public void CheckKeynoteExists(List<string> excelKeynoteValues)
        {
            if(IsKeynoteOk)
            {
                if (!excelKeynoteValues.Contains(Keynote))
                {
                    IsKeynoteOk = false;
                    KeynoteStatus = KeynoteStatus.NotListed;
                    Hash = Instance.Id.IntegerValue.ToString() + CategoryName + InstanceType.Name + Keynote + Enum.GetName(typeof(KeynoteStatus), KeynoteStatus);
                }
            }
        }

        public void TrySplitKeynote()
        {
            IsKeynoteOk = false;
            if (string.IsNullOrEmpty(Keynote))
            {
                KeynoteStatus = KeynoteStatus.Empty;
                return;
            }
                string[] split = Keynote.Split('.');
            if (split.Length == 5)
            {
                Key0 = split[0];
                Key1 = split[1];
                Key2 = split[2];
                Key3 = split[3];
                Key4 = split[4];
                IsKeynoteOk = true;
                KeynoteStatus = KeynoteStatus.OK;
            }
            else
            {
                KeynoteStatus = KeynoteStatus.IncorrectFormat;
            }
        }

        private void GetHash()
        {
            Hash = Instance.Id.IntegerValue.ToString() + CategoryName +InstanceType.Name + Keynote + Enum.GetName(typeof(KeynoteStatus), KeynoteStatus);

        }
        private void GetDimensions()
        {
            Parameter a = Instance.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if(a != null)
            {
                Area = UnitUtils.ConvertFromInternalUnits(a.AsDouble(), UnitTypeId.SquareMeters);
            }
            Parameter v = Instance.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (v != null)
            {
                Volume = UnitUtils.ConvertFromInternalUnits(v.AsDouble(), UnitTypeId.CubicMeters);
            }
            else
            {
                Parameter rv = Instance.get_Parameter(BuiltInParameter.REINFORCEMENT_VOLUME);
                if(rv != null) Volume = UnitUtils.ConvertFromInternalUnits(rv.AsDouble(), UnitTypeId.CubicMeters);
            }

            if(Instance.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Rebar)
            {
                //REBAR_ELEM_TOTAL_LENGTH
                Parameter rebLength = Instance.get_Parameter(BuiltInParameter.REBAR_ELEM_TOTAL_LENGTH);
                if (rebLength != null) Length = UnitUtils.ConvertFromInternalUnits(rebLength.AsDouble(), UnitTypeId.Meters);
            }
            else
            {
                Parameter l = Instance.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM);
                if (l != null) Length = UnitUtils.ConvertFromInternalUnits(l.AsDouble(), UnitTypeId.Meters);
            }
        }

        private bool GetInstanceType()
        {
            ElementId typeId = Instance.GetTypeId();
            InstanceType = _doc.GetElement(typeId);
            if(InstanceType == null)
            {
                return false;
            }
            else
            {
                Parameter d = InstanceType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                if (d != null && d.AsString() != null) Description = d.AsString();
                return true;
            }
        }

        private bool GetKeynote()
        {
            Parameter p = InstanceType.get_Parameter(BuiltInParameter.KEYNOTE_PARAM);
            if (p != null)
            {
                Keynote = p.AsString();
                return true;
            }
            else
            {
                return false;
            }
        }

        private void GetCategoryName()
        {
            CategoryName = Instance?.Category.Name;
        }
    }
}
