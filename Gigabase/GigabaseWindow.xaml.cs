using System.Windows.Threading;
using CefSharp;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Windows;
using TeslaRevitTools;
using System.ComponentModel;

namespace NavisworksGigabaseTsla
{
    /// <summary>
    /// Interaction logic for TeslaDockPanel.xaml
    /// </summary>
    public partial class GigabaseWindow : Window
    {
        public GigabaseViewModel GigabaseViewModel { get; set; }

        public GigabaseWindow(GigabaseViewModel gigabaseViewModel)
        {
            GigabaseViewModel = gigabaseViewModel;
            DataContext = GigabaseViewModel;
            //SetOwner();
            InitializeComponent();
            GigabaseViewModel.Browser = Browser;
            Loaded += GigabaseWindow_Loaded;
        }

        private void GigabaseWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.DoWork += Worker_DoWork;
            GigabaseViewModel.Worker = worker;
            worker.RunWorkerAsync();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            GigabaseViewModel.GigabaseAction = TeslaRevitTools.Gigabase.GigabaseAction.Collect;
            GigabaseViewModel.TheEvent.Raise();
            GigabaseViewModel.SignalEvent.WaitOne();
            GigabaseViewModel.SignalEvent.Reset();
            //GigabaseViewModel.CollectModelItems();
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            LoadingProgressBar.Value = (double)e.ProgressPercentage;
            ProgressState.Text = e.UserState.ToString();
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SubscribeEvents();
        }

        public void SubscribeEvents()
        {
            Browser.JavascriptMessageReceived += Browser_JavascriptMessageReceived;
            Browser.ConsoleMessage += Browser_ConsoleMessage;
            GigabaseViewModel.LogInfo += "Subscribing events..." + Environment.NewLine;
        }

        public void UnsubscribeEvents()
        {
            Browser.JavascriptMessageReceived -= Browser_JavascriptMessageReceived;
            Browser.ConsoleMessage -= Browser_ConsoleMessage;
            GigabaseViewModel.LogInfo += "Unubscribing events..." + Environment.NewLine;
        }

        private void SetOwner()
        {
            WindowHandleSearch windowHandleSearch = WindowHandleSearch.MainWindowHandle;
            windowHandleSearch.SetAsOwner(this);
        }

        private void Browser_ConsoleMessage(object sender, ConsoleMessageEventArgs e)
        {
            string logMessage = e.Message;
            //GigabaseViewModel.LogInfo += "Console message received:" + Environment.NewLine;
            GigabaseViewModel.LogInfo += logMessage + Environment.NewLine;
        }

        private void Browser_JavascriptMessageReceived(object sender, JavascriptMessageReceivedEventArgs e)
        {

            GigabaseViewModel.LogInfo += "CefSharp message received:" + Environment.NewLine;

            IDictionary<string, object> propertyValues = e.Message as ExpandoObject;
            object action = propertyValues["ActionType"];
            if (action != null && action as string != null)
            {
                string actionString = action as string;
                GigabaseViewModel.LogInfo += "Action type: " + actionString + Environment.NewLine;

                if (actionString == "SelectInNavis")
                {
                    UnsubscribeEvents();
                    IDictionary<string, object> data = propertyValues["Data"] as ExpandoObject;
                    GigabaseViewModel.MarksSelection = data["SelectedElements"] as List<object>;
                    GigabaseViewModel.GigabaseAction = TeslaRevitTools.Gigabase.GigabaseAction.Select;

                    GigabaseViewModel.TheEvent.Raise();
                    GigabaseViewModel.SignalEvent.WaitOne();
                    GigabaseViewModel.SignalEvent.Reset();
                    SubscribeEvents();
                    //GigabaseViewModel.SelectByMarkInCustomParameters();
                }
                if (actionString == "ImportSelected")
                {
                    GigabaseViewModel.SendSelectedElements();
                }
                if (actionString == "ImportComplete")
                {
                    MessageBox.Show("Import complete", "Import Info");
                }
            }
        }

        private void GoToAddressBtnClick(object sender, RoutedEventArgs e)
        {
            Browser.Address = AddressBox.Text;
        }

        private void LogBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            LogBox.ScrollToEnd();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            GigabaseViewModel.GigabaseWindow = null;
            Browser.Dispose();
        }

        private void ExtendColumnBtnClick(object sender, RoutedEventArgs e)
        {
            if(ExtendableColumn.Width == new GridLength(8, GridUnitType.Pixel))
            {
                ExtendableColumn.Width = new GridLength(1, GridUnitType.Star);
            }
            else
            {
                ExtendableColumn.Width = new GridLength(8, GridUnitType.Pixel);
            }
        }

        private void SelectionChangedChckBox_Checked(object sender, RoutedEventArgs e)
        {
            GigabaseViewModel.SubscribeSelectionChanged();
        }

        private void SelectionChangedChckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            GigabaseViewModel.UnsubscribeSelectionChanged();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GigabaseViewModel.UnsubscribeSelectionChanged();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
