using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace oDataToXls.Models
{
    public class oData
    {
        [JsonProperty("odata.metadata")]
        public string metadata { get; set; }
        public List<oDataValues> value { get; set; }
    }

    public class oDataValues
    {
        public string Key { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
    }
}