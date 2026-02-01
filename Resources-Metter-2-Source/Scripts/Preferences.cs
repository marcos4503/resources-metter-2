using Newtonsoft.Json;
using System.IO;

namespace Resources_Metter_2.Scripts
{
    /*
     * This class manage the load and save of program settings
    */

    public class Preferences
    {
        //Classes of script
        public class LoadedData
        {
            //*** Data to be saved ***//

            public SaveInfo[] saveInfo = new SaveInfo[0];

            public int windowPosX = 0;
            public int windowPosY = 0;
            public bool autoRestoreWindows = false;
            public int cpuFanPwmSlot = -1;
            public int cpuOptPwmSlot = -1;
            public int cpuPumpPwmSlot = -1;
            public int cpuPumpRpmWarn = -1;
            public bool autoStartWithSystem = false;
        }

        //Public variables
        public LoadedData loadedData = null;

        //Core methods

        public Preferences()
        {
            //Check if save file exists
            bool saveExists = File.Exists((Directory.GetCurrentDirectory() + @"/Content/prefs.json"));

            //If have a save file, load it
            if (saveExists == true)
                Load();
            //If a save file don't exists, create it
            if (saveExists == false)
                Save();
        }

        private void Load()
        {
            //Load the data
            string loadedDataString = File.ReadAllText((Directory.GetCurrentDirectory() + @"/Content/prefs.json"));

            //Convert it to a loaded data object
            loadedData = JsonConvert.DeserializeObject<LoadedData>(loadedDataString);
        }

        //Public methods

        public void Save()
        {
            //If the loaded data is null, create one
            if (loadedData == null)
                loadedData = new LoadedData();

            //Save the data
            File.WriteAllText((Directory.GetCurrentDirectory() + @"/Content/prefs.json"), JsonConvert.SerializeObject(loadedData));

            //Load the data to update loaded data
            Load();
        }

        /*
         * Auxiliar classes
         * 
         * Classes that are objects that will be used, only to organize data inside 
         * "LoadedData" object in the saves.
        */

        public class SaveInfo
        {
            public string key = "";
            public string value = "";
        }
    }
}
