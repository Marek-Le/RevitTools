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
    /// Interaction logic for FindReplaceWindow.xaml
    /// </summary>
    public partial class FindReplaceWindow : Window
    {
        public bool WindowResult { get; set; } = false;
        public FindReplaceWindow(Window owner)
        {
            Owner = owner;
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
