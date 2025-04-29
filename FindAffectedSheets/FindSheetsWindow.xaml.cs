using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TeslaRevitTools.FindAffectedSheets
{
    /// <summary>
    /// Interaction logic for FindSheetsWindow.xaml
    /// </summary>
    public partial class FindSheetsWindow : Window
    {
        public FindSheetsViewModel FindSheetsViewModel { get; set; }

        public FindSheetsWindow(FindSheetsViewModel findSheetsViewModel)
        {
            FindSheetsViewModel = findSheetsViewModel;
            DataContext = FindSheetsViewModel;
            SetAsOwner(findSheetsViewModel.IntPtr);
            InitializeComponent();
        }

        private void SetAsOwner(IntPtr handle)
        {
            new WindowInteropHelper(this) { Owner = handle };
        }

        private void CheckButton_Click(object sender, RoutedEventArgs e)
        {
            FindSheetsViewModel.FoundSheets.Clear();
            FindSheetsViewModel.FoundViews.Clear();
            CheckButton.Visibility = System.Windows.Visibility.Collapsed;
            ProgressBar.Visibility = System.Windows.Visibility.Visible;
            ProgressState.Visibility = System.Windows.Visibility.Visible;

            BackgroundWorker worker = new BackgroundWorker();
            FindSheetsViewModel.Worker = worker;
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerAsync();
        }

        private void Worker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            for (int i = 0; i < FindSheetsViewModel.AffectedViews.Count; i++)
            {
                View v = FindSheetsViewModel.AffectedViews[i];
                if (!FindSheetsViewModel.FoundViews.Select(x => x.Id).Contains(v.Id))
                {
                    FindSheetsViewModel.FoundViews.Add(v);
                }
            }

            ProgressBar.Value = (double)e.ProgressPercentage;
            ProgressState.Text = e.UserState.ToString();
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            for (int i = 0; i < FindSheetsViewModel.AffectedSheets.Count; i++)
            {
                ViewSheet vs = FindSheetsViewModel.AffectedSheets[i];
                if (!FindSheetsViewModel.FoundSheets.Select(x => x.Id).Contains(vs.Id))
                {
                    FindSheetsViewModel.FoundSheets.Add(vs);
                }
            }
            ProgressBar.Value = 100;
            ProgressState.Text = "Checking finished!";
            CheckButton.Visibility = System.Windows.Visibility.Visible;
            ProgressBar.Visibility = System.Windows.Visibility.Collapsed;
            ProgressState.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

            FindSheetsViewModel.Action = FindSheetsAction.FindSheets;
            FindSheetsViewModel.SheetSelectionEvent.Raise();
            FindSheetsViewModel.SignalEvent.WaitOne();
            FindSheetsViewModel.SignalEvent.Reset();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            FindSheetsViewModel.SelectedElements.Clear();
            FindSheetsViewModel.SelectionChangedSubscriber.Raise();
        }

        private void AffectedSheetsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var selectedSheet = AffectedSheetsGrid.SelectedItem as ViewSheet;
            if(selectedSheet != null) 
            { 
                FindSheetsViewModel.SelectedView = selectedSheet;
                FindSheetsViewModel.Action = FindSheetsAction.SelectView;
                FindSheetsViewModel.SheetSelectionEvent.Raise();
            }
        }

        private void AffectedViewsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var selectedView = AffectedViewsGrid.SelectedItem as View;
            if (selectedView != null)
            {
                FindSheetsViewModel.SelectedView = selectedView;
                FindSheetsViewModel.Action = FindSheetsAction.SelectView;
                FindSheetsViewModel.SheetSelectionEvent.Raise();
            }
        }
    }
}
