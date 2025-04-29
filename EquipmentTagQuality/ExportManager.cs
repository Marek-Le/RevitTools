using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using TeslaRevitTools.EquipmentTagQuality.GigabaseModels;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class ExportManager
    {
        public List<EquipmentType> eqTypes;
        public List<SystemType> sysTypes;
        public List<EquipmentLocation> locations;

        #region GIGABASE DEV Endpoints:
        string equipment_category_url = "https://gigamap-dev.tesla.com/backend/db/equipment_category";
        string system_type_url = "https://gigamap-dev.tesla.com/backend/db/system";
        string location_url = "https://gigamap-dev.tesla.com/backend/db/location";
        #endregion

        public ExportManager()
        {
            eqTypes = GetEquipmentTypes();
            sysTypes = GetSystemTypes();
            locations = GetLocations();
        }

        public List<EquipmentLocation> GetLocations()
        {
            List<EquipmentLocation> locations = new List<EquipmentLocation>();
            ProcessGigabaseData(ref locations, location_url, out string ex);
            return locations;
        }

        public List<SystemType> GetSystemTypes()
        {
            List<SystemType> systemTypes = new List<SystemType>();
            ProcessGigabaseData(ref systemTypes, system_type_url, out string ex);

            return systemTypes;
        }

        public List<EquipmentType> GetEquipmentTypes()
        {
            List<EquipmentType> equipmentTypes = new List<EquipmentType>();
            ProcessGigabaseData(ref equipmentTypes, equipment_category_url, out string ex);
            return equipmentTypes;
        }

        private bool ProcessGigabaseData<T>(ref List<T> listValues, string url, out string exception)
        {
            try
            {
                exception = "";
                HttpClient httpClient = new HttpClient();
                var result = httpClient.GetAsync(url).Result;
                var jsonString = result.Content.ReadAsStringAsync().Result;
                listValues = JsonConvert.DeserializeObject<List<T>>(jsonString);
                return true;
            }
            catch (Exception ex)
            {
                exception = ex.ToString();
                return false;
            }
        }

        private void CalculateLevel()
        {
            //Level 1 = 39400mm  -  0,0

            //1F = 0  for all shops
            //1M DU = 7.5 m
            //1M BW = 7.56 m
            //1M SB = 6.0 m
            //1M CB = 3.35 m
            //2F DU = 6.0 m
            //2F BW PT GA = 8.5 m
            //2F CP = 3.0 m
            //2F PL = 8.6 m
            //2F SB = 7.5 m
            //2F WW = 4.75
            //2F CR = 7.1 m
            //2F CB = 8.35

            Dictionary<string, List<ShopLevel>> shopLevels = new Dictionary<string, List<ShopLevel>>();
            shopLevels.Add("DU", new List<ShopLevel>() { new ShopLevel() { LevelName = "1F", AbsoluteElevation = 39400, RelativeElevation = 0 } });
        }

        class ShopLevel
        {
            public string LevelName { get; set; }
            public double AbsoluteElevation { get; set; }
            public double RelativeElevation { get; set; }

        }
    }
}
