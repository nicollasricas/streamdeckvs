using Newtonsoft.Json;

namespace VisualStudioStreamDeck
{
    public class ExecuteCommandSettings : KeySettings
    {
        [JsonProperty("command")]
        public string Command { get; set; }
    }
}
