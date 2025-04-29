using Microsoft.WindowsAPICodePack.Dialogs;
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

namespace TeslaRevitTools.FileCleaner
{
    /// <summary>
    /// Interaction logic for FileCleanerWindow.xaml
    /// </summary>
    public partial class FileCleanerWindow : Window
    {
        public FileCleanerModel FileCleanerModel { get; set; }

        public FileCleanerWindow(FileCleanerModel fileCleanerModel)
        {
            SetOwner();
            FileCleanerModel = fileCleanerModel;
            this.DataContext = FileCleanerModel;
            InitializeComponent();
            //CreateTestLinks();
            CreateComBoxItems();
            SetLabelBrush();
        }

        public void SetLabelBrush()
        {
            if(FileCleanerModel.StatusLabel == "Activated")
            {
                StatusLabel.Foreground = Brushes.Green;
            }
            else
            {
                StatusLabel.Foreground = Brushes.Red;
            }
        }

        private void SetOwner()
        {
            WindowHandleSearch windowHandleSearch = WindowHandleSearch.MainWindowHandle;
            windowHandleSearch.SetAsOwner(this);
        }

        //public void CreateTestLinks()
        //{
        //    List<string> links = new List<string>();
        //    links.Add("first path");
        //    links.Add("second path");
        //    LinksBox.ItemsSource = links;
        //}

        private void LinksBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Delete)
            {
                DeleteSelectedLink();
                e.Handled = true;
            }
        }

        private bool DeleteSelectedLink()
        {
            bool result = false;
            if(LinksBox.SelectedItems.Count == 1)
            {
                result = true;
                (LinksBox.ItemsSource as List<string>).Remove(LinksBox.SelectedItem as string); 
                LinksBox.Items.Refresh();
            }
            return result;
        }

        private void CreateComBoxItems()
        {
            List<string> days = new List<string>();
            for (int i = 0; i < 30; i++) days.Add((i + 1).ToString());
            ComBoxDays.ItemsSource = days;
            //ComBoxDays.SelectedIndex = 29;
        }

        private void LinksBox_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                try
                {
                    if (System.IO.Directory.Exists(dialog.FileName))
                    {
                        string[] files = System.IO.Directory.EnumerateFiles(dialog.FileName, "*", System.IO.SearchOption.AllDirectories).Take(101).ToArray();
                        if (files.Length > 100)
                        {
                            var result = MessageBox.Show("There is more than 100 files in this directory and its subdirectories! Are you sure???", "Warning", MessageBoxButton.YesNo);
                            if (result == MessageBoxResult.Yes)
                            {
                                FileCleanerModel.DirectoryPaths.Add(dialog.FileName);
                            }
                        }
                        else
                        {
                            FileCleanerModel.DirectoryPaths.Add(dialog.FileName);
                        }
                    }
                    LinksBox.Items.Refresh();
                }
                catch (Exception)
                {
                    MessageBox.Show("We can not add this directory to list!");
                }
            }
        }

        private void BtnClick_Deactivate(object sender, RoutedEventArgs e)
        {
            FileCleanerModel.IsActive = false;
            FileCleanerModel.StatusLabel = "Deactivated";
            StatusLabel.Foreground = Brushes.Red;
            FileCleanerModel.WriteJsonSettings();
        }

        private void BtnClick_SaveAndActivate(object sender, RoutedEventArgs e)
        {
            FileCleanerModel.DaysAllowed = ComBoxDays.SelectedIndex + 1;
            FileCleanerModel.IsActive = true;
            FileCleanerModel.StatusLabel = "Activated";
            StatusLabel.Foreground = Brushes.Green;
            FileCleanerModel.WriteJsonSettings();
        }
    }
}
