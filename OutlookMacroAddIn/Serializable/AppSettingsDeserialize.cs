using Newtonsoft.Json;
using OutlookMacroAddIn.Serializable.Entity;
using System.IO;

namespace OutlookMacroAddIn.Serializable
{
    public class AppSettingsDeserialize
    {
        private readonly string jsonPatch;

        public AppSettingsDeserialize(string jsonPatch)
        {
            this.jsonPatch = jsonPatch;
        }

        public AppSettings GetSettingsModels()
        {
            JsonSerializer serializer = new JsonSerializer();
            using (StreamReader sw = new StreamReader(jsonPatch))
            {
                using (JsonReader reader = new JsonTextReader(sw))
                {
                    return serializer.Deserialize<AppSettings>(reader);
                }
            }
        }
    }
}
