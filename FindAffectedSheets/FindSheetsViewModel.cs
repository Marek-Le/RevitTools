using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Threading;
using TeslaRevitTools.ITwoExport;
using TeslaRevitTools.RvtExtEvents;

namespace TeslaRevitTools.FindAffectedSheets
{
    public class FindSheetsViewModel : INotifyPropertyChanged
    {
        public Thread WindowThread { get; set; }
        public FindSheetsAction Action { get; set; }
        public ManualResetEvent SignalEvent { get; set; }
        public BackgroundWorker Worker { get; set; }
        public ExternalEvent SheetSelectionEvent { get; set; }
        public ExternalEvent SelectionChangedSubscriber { get; set; }
        public FindSheetsWindow TheWindow { get; set; }
        private List<SelectedElement> _selectedElements;
        public List<SelectedElement> SelectedElements
        {
            get { return _selectedElements; }
            set
            {
                if(_selectedElements?.Count != value.Count) 
                {
                    _selectedElements = value;
                    OnPropertyChanged(nameof(SelectedElements));
                }
            }
        }

        public ObservableCollection<ViewSheet> FoundSheets { get; set; }
        public List<ViewSheet> AffectedSheets { get; set; }
        public ObservableCollection<View> FoundViews { get; set; }
        public List<View> AffectedViews { get; set; }
        public Document ActiveDocument { get; set; }
        public IntPtr IntPtr { get; set; }
        public View SelectedView { get; set; }

        public FindSheetsViewModel()
        {
            FindSheetsEventHandler findSheetsEventHandler = new FindSheetsEventHandler();
            SheetSelectionEvent = ExternalEvent.Create(findSheetsEventHandler);
            SignalEvent = new ManualResetEvent(false);
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void FindAffectedSheetViews()
        {
            try
            {
                List<ViewSheet> allSheets = new FilteredElementCollector(ActiveDocument).OfClass(typeof(ViewSheet)).Select(e => e as ViewSheet).ToList();
                List<ElementId> viewIds = new List<ElementId>();
                foreach (ViewSheet vs in allSheets)
                {
                    foreach (ElementId id in vs.GetAllPlacedViews())
                    {
                        if (!viewIds.Contains(id)) viewIds.Add(id);
                    }
                }

                

                List<View> viewOnSheets = viewIds.Select(e => ActiveDocument.GetElement(e) as View).ToList();

                List<View> viewsContainingElementBbox = new List<View>();


                foreach (var view in viewOnSheets)
                {
                    bool isInsideViewCheck = false;
                    CheckedView checkedView = new CheckedView(view);
                    foreach (SelectedElement se in SelectedElements)
                    {
                        if (checkedView.CropBoxContains(se.Element, ActiveDocument))
                        {
                            isInsideViewCheck = true;
                            break;
                        }
                    }
                    if (isInsideViewCheck) viewsContainingElementBbox.Add(view);
                }

                AffectedViews = new List<View>();
                AffectedSheets = new List<ViewSheet>();

                int progress = 0;
                int totalItems = viewsContainingElementBbox.Count;
                double percentage = ((double)progress / (double)totalItems) * 100;
                Worker.ReportProgress((int)percentage, string.Format("{0} of {1} views checked", progress, totalItems));

                foreach (var view in viewsContainingElementBbox)
                {
                    if (view.Id.IntegerValue == 3606959)
                    {

                    }
                    ICollection<ElementId> ids = new FilteredElementCollector(ActiveDocument, view.Id).WhereElementIsNotElementType().Select(e => e.Id).ToList();
                    foreach (SelectedElement se in SelectedElements)
                    {
                        if (ids.Select(e => e.IntegerValue).Contains(se.ElementId))
                        {
                            if (AffectedViews.Select(e => e.Id.IntegerValue).Contains(se.ElementId)) continue;
                            AffectedViews.Add(view);
                        }
                    }
                    percentage = ((double)progress / (double)totalItems) * 100;
                    Worker.ReportProgress((int)percentage, string.Format("{0} of {1} views checked", progress, totalItems));
                    progress++;
                }

                List<ViewSheet> sheets = new FilteredElementCollector(ActiveDocument).OfClass(typeof(ViewSheet)).Select(e => e as ViewSheet).ToList();
                foreach (ViewSheet vs in sheets)
                {
                    if (vs.GetAllPlacedViews().Any(x => AffectedViews.Any(y => y.Id.IntegerValue == x.IntegerValue)))
                    {
                        AffectedSheets.Add(vs);
                    }
                }
                SignalEvent.Set();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
                SignalEvent.Set();
            }
        }

        public void DisplayWindow(UIApplication application)
        {
            SelectedElements = new List<SelectedElement>();
            FoundSheets = new ObservableCollection<ViewSheet>();
            FoundViews = new ObservableCollection<View>();
            ActiveDocument = application.ActiveUIDocument.Document;
            IntPtr = application.MainWindowHandle;
            SelectionChangedSubscriber.Raise();
            WindowThread = new Thread(delegate ()
            {
                TheWindow = new FindSheetsWindow(this);
                TheWindow.Show();
                Dispatcher.Run();
            });
            WindowThread.SetApartmentState(ApartmentState.STA);
            WindowThread.IsBackground = true;
            WindowThread.Start();
        }

        public void GetSelectedElements(List<ElementId> selectedIds)
        {
            SelectedElements.Clear();
            foreach (ElementId id in selectedIds)
            {
                Element instance = ActiveDocument.GetElement(id);
                if (instance != null)
                {
                    if(instance.Category != null && instance.Category.CategoryType == CategoryType.Model) 
                    {
                        SelectedElement selectedElement = new SelectedElement()
                        {
                            ElementId = instance.Id.IntegerValue,
                            TypeName = instance.GetType().Name,
                            CategoryName = instance.Category.Name,
                            Element = instance,
                        };

                        SelectedElements.Add(selectedElement);
                    }
                }
            }
            Dispatcher.FromThread(WindowThread).Invoke(new Action(() => { 
                TheWindow.SelectedElementsGrid.Items.Refresh();
                if(SelectedElements.Count > 0)
                {
                    TheWindow.CheckButton.IsEnabled = true;
                }
                else
                {
                    TheWindow.CheckButton.IsEnabled = false;
                }
            }));
        }
    }
}
