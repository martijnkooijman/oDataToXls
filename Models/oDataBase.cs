using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace oDataToXls.Models
{
    public class oDataBase
    {
        [JsonProperty("odata.metadata")]
        public string metadata { get; set; }
        public List<oDataBaseValue> value { get; set; }
    }

    public class oDataBaseValue
    {
        public string name { get; set; }
        public string url { get; set; }
    }
}