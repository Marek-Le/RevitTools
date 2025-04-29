using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TeslaRevitTools.BatchPrint;
using TeslaRevitTools.RibbonControl;
using System.Reflection;
using TeslaRevitTools.Pathfinder;
using TeslaRevitTools.GenerateSheets;
using TeslaRevitTools.RvtExtEvents;
using TeslaRevitTools.VisTableImport;
using TeslaRevitTools.EquipmentTagQuality;
using NavisworksGigabaseTsla;
using System.Windows.Threading;
using System;
using TeslaRevitTools.GridReference;
using TeslaRevitTools.Openings;
using TeslaRevitTools.ITwoExport;
using TeslaRevitTools.FindAffectedSheets;
using TeslaRevitTools.ConvoidOpenings;

namespace TeslaRevitTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        #region Configuration and assembly version

#if DEBUG2020 || RELEASE2020
        public const string AssemblyYear = "2020";
#elif DEBUG2023 || RELEASE2023
        public const string AssemblyYear = "2023";
#endif

        //increment on every release build
        public const string AssemblyMinorVersion = "5"; //new functions
        public const string AssemblyRevisionVersion = "7"; //significant fixes and improvements
        public const string AssemblyBuildVersion = "0"; //small fixes and improvements

        public static string PlugInVersion
        {
            get
            {
                return string.Format("{0}.{1}.{2}.{3}",
                                AssemblyYear,
                                AssemblyMinorVersion,
                                AssemblyRevisionVersion,
                                AssemblyBuildVersion);
            }
        }
        #endregion

        public GridRefViewModel GridRefViewModel { get; set; }
        public BatchPrintViewModel BatchPrintViewModel { get; set; }

        public PathfinderViewModel PathfinderViewModel { get; set; }
        public GenerateSheetsViewModel GenerateSheetsViewModel { get; set;}
        public AdvancedElementSelector.SelectorViewModel SelectorViewModel { get; set; }
        //public VisTableViewModel VisTableViewModel { get; set; }
        public VisTableAppModel VisTableAppModel { get; set; }
        public TagQualityViewModel TagQualityViewModel { get; set; }
        public GigabaseViewModel GigabaseViewModel { get; set; }
        public OpeningsViewModel OpeningsViewModel { get; set; }
        public ExportViewModel ExportViewModel { get; set; }
        public FindSheetsViewModel FindSheetsViewModel { get; set; }
        public ConvoidViewModel ConvoidViewModel { get; set; }

        public static UIApplication UIApp;
        public static UIControlledApplication UIContApp;
        internal static App _app = null;
        public static App Instance
        {
            get { return _app; }
        }
        public Result OnStartup(UIControlledApplication app)
        {
            
            _app = this;           
            UIContApp = app;
            UIApp = GetUiApplication();      
            BatchPrintViewModel = new BatchPrintViewModel();
            FindSheetsViewModel = new FindSheetsViewModel();
            TeslaRibbonBuilder teslaRibbonBuilder = new TeslaRibbonBuilder(app);
            GenerateSheetsViewModel = new GenerateSheetsViewModel();
            CreateViewsExEventHandler createViewsExEventHandler = new CreateViewsExEventHandler();
            ExternalEvent createViewsExEvent = ExternalEvent.Create(createViewsExEventHandler);
            GenerateSheetsViewModel.TheEvent = createViewsExEvent;
            VisTableAppModel = new VisTableAppModel();
            VisTableAppModel.SetExternalEvent();
            GridRefViewModel = new GridRefViewModel();
            OpeningsViewModel = new OpeningsViewModel();
            iTwoExportExEvent iTwoExportExEvent = new iTwoExportExEvent();
            ExportViewModel = ExportViewModel.Initialize(ExternalEvent.Create(iTwoExportExEvent));
            ConvoidViewModel = new ConvoidViewModel();
            ConvoidViewModel.DownloadAndExtractData();
            ConvoidOpeningsEvent convoidOpeningsEvent = new ConvoidOpeningsEvent();
            ConvoidViewModel.TheEvent = ExternalEvent.Create(convoidOpeningsEvent);
            ConvoidViewModel.CheckAppData();
            app.ControlledApplication.DocumentOpened += ConvoidViewModel.ControlledApplication_DocumentOpened;
            //app.ControlledApplication.LinkedResourceOpened += ControlledApplication_LinkedResourceOpened;
            DockPanelModel dockPanel = new DockPanelModel();
            dockPanel.RegisterDockPanel(app, ConvoidViewModel);
            teslaRibbonBuilder.FindCreateTeslaRibbonTab();
            teslaRibbonBuilder.AddProjectToolsPanel();
            teslaRibbonBuilder.AddBatchPrintButton();
            teslaRibbonBuilder.AddCleanFilesButton();
            teslaRibbonBuilder.AddFamilyRenameButton();
            teslaRibbonBuilder.AddTemplateOverridesButton();
            teslaRibbonBuilder.AddSettingsButton();
            //teslaRibbonBuilder.AddPathFinderButton();
            teslaRibbonBuilder.AddGenerateSheetsButton();
            teslaRibbonBuilder.AddFindSheetsBtn();
            //teslaRibbonBuilder.AddModelOptimiserButton();
            teslaRibbonBuilder.AddViewTemplatesManager();
            teslaRibbonBuilder.AddTagQualityTool();
            teslaRibbonBuilder.AddGigaBaseTool();
            teslaRibbonBuilder.AddGridRefTool();
            teslaRibbonBuilder.AddOpeningsTool();
#if DEBUG2023 || RELEASE2023
            teslaRibbonBuilder.AddVisTableImportButton();
            teslaRibbonBuilder.AddITwoExport();
#endif
            return Result.Succeeded;
        }

        private void ControlledApplication_LinkedResourceOpened(object sender, Autodesk.Revit.DB.Events.LinkedResourceOpenedEventArgs e)
        {
            //TaskDialog.Show("Info", "Opened linked resource");
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            if (GigabaseViewModel != null && GigabaseViewModel.WindowThread != null)
            {
                try
                {
                    if (GigabaseViewModel.GigabaseWindow != null)
                    {
                        Dispatcher.FromThread(GigabaseViewModel.WindowThread).Invoke(new Action(() => { 
                            GigabaseViewModel.GigabaseWindow.Close();
                            GigabaseViewModel.Browser.Dispose();
                        }));
                    }
                    CefSharp.Cef.Shutdown();
                    GigabaseViewModel.WindowThread.Abort();
                }
                catch (Exception ex)
                {

                }
            }
            FileCleaner.FileCleanerModel.DeleteFiles();
            return Result.Succeeded;
        }

        private static UIApplication GetUiApplication()
        {
            var versionNumber = UIContApp.ControlledApplication.VersionNumber;
            var fieldName = string.Empty;
            switch (versionNumber)
            {
                case "2020":
                    fieldName = "m_uiapplication";
                    break;
                case "2023":
                    fieldName = "m_uiapplication";
                    break;
            }
            var fieldInfo = UIContApp.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
            var uiApplication = (UIApplication)fieldInfo?.GetValue(UIContApp);
            return uiApplication;
        }
    }
}
