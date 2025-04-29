using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.ITwoExport
{
    public class CategoryModel : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }
        public string Name { get; set; }
        public BuiltInCategory BuiltInCategory { get; set; }
        private float _density;
        public float Density
        {
            get { return _density; } 
            set 
            {          
                if (_density != value)
                {
                    _density = value; 
                    OnPropertyChanged(nameof(Density));
                }
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
