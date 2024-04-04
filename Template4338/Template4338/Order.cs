using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Template4338
{
    public partial class Order
    {
        [JsonIgnore]
        public int Id { get; set; }

        [JsonPropertyName("CodeStaff")]
        public string CodeStaff { get; set; }

        [JsonPropertyName("Position")]
        public string Position { get; set; }

        [JsonPropertyName("FullName")]
        public string FullName { get; set; }

        [JsonPropertyName("Log")]
        public string Log { get; set; }

        [JsonPropertyName("Password")]
        public string Password { get; set; }

        [JsonPropertyName("LastEnter")]
        public string LastEnter { get; set; }

        [JsonPropertyName("TypeEnter")]
        public string TypeEnter { get; set; }

       
    }
}
