using System.Collections.Generic;
using Newtonsoft.Json;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class GraphDataSet<T>
    {
        [JsonProperty(PropertyName ="value")]
        public List<T> Value;
    }
}
