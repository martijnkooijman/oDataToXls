using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace oDataToXls.Models
{
    public class DataProperties
    {
        [JsonProperty("odata.metadata")]
        public string metadata { get; set; }
        public List<DataPropertiesValue> value { get; set; }
    }

    public class DataPropertiesValue
    {
        [JsonProperty("odata.type")]
        public string oDataType { get; set; }

        public int? ID { get; set; }
        public int? Position { get; set; }

        public int? ParentID { get; set; }
        public string Type { get; set; }
        public string Key { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
    }
}