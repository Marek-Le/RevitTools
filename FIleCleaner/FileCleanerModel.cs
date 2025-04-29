using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.ComponentModel;

namespace TeslaRevitTools.FileCleaner
{
    public class FileCleanerModel : INotifyPropertyChanged
    {
        private List<string> _directoryPaths;
        public List<string> DirectoryPaths
        {
            get => _directoryPaths;
            set
            {
                if (_directoryPaths != value)
                {
                    _directoryPaths = value;
                    OnPropertyChanged(nameof(DirectoryPaths));
                }
            }
        }
        public bool IsActive { get; set; }
        public bool DeleteDirectories { get; set; }
        public int DaysAllowed { get; set; }
        public int DaysAllowedIndex { get; set; }

        private string _statusLabel;
        public string StatusLabel
        {
            get => _statusLabel;
            set
            {
                if (_statusLabel != value)
                {
                    _statusLabel = value;
                    OnPropertyChanged(nameof(StatusLabel));
                }
            }
        }

        string _jsonFilePath;

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private FileCleanerModel()
        {

        }

        public static FileCleanerModel Initialize()
        {
            FileCleanerModel result = new FileCleanerModel();
            result.GetJsonSettingsFilePath();
            result.ReadJsonSettings();
            return result;
        }

        public void DisplayWindow()
        {
            FileCleanerWindow window = new FileCleanerWindow(this);
            window.ComBoxDays.SelectedIndex = DaysAllowedIndex;
            window.ShowDialog();
        }

        private void GetJsonSettingsFilePath()
        {
            string assemblyDirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tslaRevitToolsDirectory = Directory.GetParent(assemblyDirPath).FullName;
            _jsonFilePath = Path.Combine(tslaRevitToolsDirectory, "cleaner_settings.json");
        }

        public void ReadJsonSettings()
        {
            string jsonString = File.ReadAllText(_jsonFilePath);
            CleanerSettings cleanerSettings = JsonConvert.DeserializeObject<CleanerSettings>(jsonString);
            DirectoryPaths = cleanerSettings.DirectoryPaths;
            IsActive = cleanerSettings.IsActive;
            if(IsActive)
            {
                StatusLabel = "Activated";
            }
            else
            {
                StatusLabel = "Deactivated";
            }
            DeleteDirectories = cleanerSettings.DeleteDirectories;
            DaysAllowed = cleanerSettings.DaysAllowed;
            DaysAllowedIndex = DaysAllowed - 1;
        }

        public void WriteJsonSettings()
        {
            CleanerSettings cleanerSettings = new CleanerSettings();
            cleanerSettings.DaysAllowed = DaysAllowed;
            cleanerSettings.DeleteDirectories = DeleteDirectories;
            cleanerSettings.DirectoryPaths = DirectoryPaths;
            cleanerSettings.IsActive = IsActive;
            string jsonString = JsonConvert.SerializeObject(cleanerSettings);
            File.WriteAllText(_jsonFilePath, jsonString);
        }

        public static void DeleteFiles()
        {
            FileCleanerModel cleaner = new FileCleanerModel();
            cleaner.GetJsonSettingsFilePath();
            cleaner.ReadJsonSettings();
            if(cleaner.IsActive)
            {
                cleaner.DeleteFiles(cleaner.DirectoryPaths.ToArray(), cleaner.DaysAllowed);
                //if(cleaner.DeleteDirectories)
                //{
                //    cleaner.DeleteEmptyDirectories(cleaner.DirectoryPaths.ToArray(), cleaner.DaysAllowed);
                //}
            }
        }

        private void DeleteFiles(string[] paths, int noOfDays)
        {
            foreach (string path in paths)
            {
                if (Directory.Exists(path))
                {
                    string[] files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);
                    foreach (string file in files)
                    {
                        DateTime now = DateTime.Now;
                        DateTime modified = File.GetLastWriteTime(file);
                        double check = (now - modified).TotalDays;
                        if (check < noOfDays) continue;
                        try
                        {
                            File.Delete(file);
                            //Console.WriteLine("File deleted: " + Path.GetFileNameWithoutExtension(file));
                        }
                        catch (Exception)
                        {
                            //Console.WriteLine("Can't delete file: " + Path.GetFileNameWithoutExtension(file));
                        }
                    }
                }
            }
        }

        private void DeleteEmptyDirectories(string[] paths, int noOfDays)
        {
            foreach (string path in paths)
            {
                if(Directory.Exists(path))
                {
                    string[] directories = Directory.GetDirectories(path);
                    foreach (string directory in directories)
                    {
                        if (Directory.GetFiles(directory).Length > 0) continue;
                        DateTime now = DateTime.Now;
                        DateTime modified = Directory.GetLastWriteTime(directory);
                        double check = (now - modified).TotalDays;
                        if (check < noOfDays) continue;
                        try
                        {
                            Directory.Delete(directory);
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
            }
        }
    }
}
