using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TeslaRevitTools.ITwoExport
{
    public class ExportedMaterial : INotifyPropertyChanged
    {
        public string Hash { get; set; }
        public ElementId InstanceId { get; set; }
        public Element Instance { get; set; }
        public Material Material { get; set; }
        public string Description { get; set; } = "";
        public string CategoryName { get; set; }
        public double Area { get; set; }
        public double Volume { get; set; }
        public bool IsPaint { get; set; }
        public double Density { get; set; }


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

        private Document _doc;

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static ExportedMaterial Initialize(Element element, Material material, Document doc, bool isPaint)
        {
            ExportedMaterial result = new ExportedMaterial();
            result.IsPaint = isPaint;
            result.Instance = element;
            result.InstanceId = element.Id;
            result.Material = material;
            result._doc = doc;
            result.GetKeynote();
            result.GetCategoryName();
            result.TrySplitKeynote();
            result.GetDimensions();
            result.GetDescription();
            result.GenerateHash();
            return result;
        }

        private void GenerateHash()
        {
            Hash = Instance.Id.IntegerValue.ToString() + Material.Name + Material.Id.ToString() + Description + CategoryName + IsPaint.ToString() + Keynote + Enum.GetName(typeof(KeynoteStatus), KeynoteStatus);
        }

        public void CheckKeynoteExists(List<string> excelKeynoteValues)
        {
            if (IsKeynoteOk)
            {
                if (!excelKeynoteValues.Contains(Keynote))
                {
                    IsKeynoteOk = false;
                    KeynoteStatus = KeynoteStatus.NotListed;
                    Hash = Instance.Id.IntegerValue.ToString() + Material.Name + Material.Id.ToString() + Description + CategoryName + IsPaint.ToString() + Keynote + Enum.GetName(typeof(KeynoteStatus), KeynoteStatus);
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

        private void GetDimensions()
        {
            Area = UnitUtils.ConvertFromInternalUnits(Instance.GetMaterialArea(Material.Id, IsPaint), UnitTypeId.SquareMeters);
            Volume = UnitUtils.ConvertFromInternalUnits(Instance.GetMaterialVolume(Material.Id), UnitTypeId.CubicMeters);
        }

        private bool GetKeynote()
        {
            Parameter p = Material.get_Parameter(BuiltInParameter.KEYNOTE_PARAM);
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

        private void GetDescription()
        {
            Parameter d = Material.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (d != null && d.AsString() != null) Description = d.AsString();
        }
    }
}
