using System.Collections.Generic;
using System.Runtime.Serialization.Json;
using System.IO;
using System.Runtime.Serialization;

namespace TsvData2Excel.src
{
    [DataContract]
    public class Identifier
    {
        [DataMember(Name = "tsv")]
        public string Tsv { get; set; }
        [DataMember(Name = "xlsx")]
        public string Xlsx { get; set; }
    }

    [DataContract]
    public class Mapping
    {
        [DataMember(Name = "tsv")]
        public IList<string> Tsv { get; set; }
        [DataMember(Name = "xlsx")]
        public IList<string> Xlsx { get; set; }
        [DataMember(Name = "splitChar")]
        public string SplitChar { get; set; }
    }

    [DataContract]
    public class Config
    {
        [DataMember(Name = "splitChar")]
        public string SplitChar { get; set; }
        [DataMember(Name = "targetSheet")]
        public string TargetSheet { get; set; }
        [DataMember(Name = "identifier")]
        public Identifier Identifier { get; set; }
        [DataMember(Name = "filledColumn")]
        public IList<string> FilledColumn { get; set; }
        [DataMember(Name = "endOfColumn")]
        public string EndOfColumn { get; set; }
        [DataMember(Name = "mapping")]
        public IList<Mapping> Mapping { get; set; }
    }

    public class ConfigSerializer
    {
        public static Config Serialize(string filePath)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Config));
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                return (Config)serializer.ReadObject(fs);
            };
        }
    }
}
