using Newtonsoft.Json;

namespace StreamDeckVS
{
    public class ExecuteCommandSettings : KeySettings
    {
        [JsonProperty("command")]
        public string Command { get; set; }
    }
}
