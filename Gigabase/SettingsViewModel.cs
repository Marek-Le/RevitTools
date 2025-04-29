using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace Settings
{
    public class SettingsViewModel
    {
        public ObservableCollection<SearchParameter> SearchParameters { get; set; }

        private string _jsonSettingsPath;
        private string _gigaBaseDirectory;

        private SettingsViewModel()
        {

        }

        public static SettingsViewModel Initialize()
        {
            SettingsViewModel result = new SettingsViewModel();
            result._gigaBaseDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Autodesk\\Navisworks Manage 2023\\TslaGigabase");
            result._jsonSettingsPath = Path.Combine(result._gigaBaseDirectory, "settings.json");
            return result;
        }

        public void SerializeAndSave()
        {
            if (!Directory.Exists(_gigaBaseDirectory)) Directory.CreateDirectory(_gigaBaseDirectory);
            string jsonString = JsonConvert.SerializeObject(SearchParameters);
            File.WriteAllText(_jsonSettingsPath, jsonString);
        }

        public void GetSearchParameters()
        {
            if (!DeserializeAndLoad())
            {
                SearchParameters = new ObservableCollection<SearchParameter>();
                SearchParameter searchParameter = new SearchParameter() { ParameterName = "Mark", CategoryName = "Element", MinLength = 5 };
                if (!SearchParameters.Select(e => e.ParameterName).Contains(searchParameter.ParameterName)) SearchParameters.Add(searchParameter);
            }
            else
            {
                for (int i = SearchParameters.Count - 1; i >= 0; i--)
                {
                    SearchParameter sp = SearchParameters[i];
                    if (string.IsNullOrEmpty(sp.ParameterName) || string.IsNullOrEmpty(sp.CategoryName)) SearchParameters.RemoveAt(i);
                }
            }
        }

        private bool DeserializeAndLoad()
        {
            if (File.Exists(_jsonSettingsPath))
            {
                string jsonString = File.ReadAllText(_jsonSettingsPath);
                SearchParameters = JsonConvert.DeserializeObject<ObservableCollection<SearchParameter>>(jsonString);
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
