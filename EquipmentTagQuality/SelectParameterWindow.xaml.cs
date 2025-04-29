using DocumentFormat.OpenXml.Office.CoverPageProps;
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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TeslaRevitTools.EquipmentTagQuality
{
    /// <summary>
    /// Interaction logic for SelectParameterWindow.xaml
    /// </summary>
    public partial class SelectParameterWindow : Window
    {
        public bool WindowResult { get; set; } = false;
        public List<string> CustomParameters { get; set; }
        public SelectParameterWindow(List<string> customParameters, Window owner)
        {
            Owner = owner;
            CustomParameters = customParameters;
            DataContext = this;
            InitializeComponent();
        }

        private void BtnCancel(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnApply(object sender, RoutedEventArgs e)
        {
            WindowResult = true;
            Close();
        }
    }
}
