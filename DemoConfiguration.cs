using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AU2024_smart_parameter_updater
{
    public class DemoConfiguration
    {
        public string ExchangeFileUrn { get; set; }
        public string CollectionId { get; set; }
        public string ClassName { get; set; }
        public string Excel { get; set; }
        public string ApplicationName { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string CallBack { get; set; }
        public string HubId { get; set; }        
        public string ProjectId { get; set; }
        //can use the REST api to read it
        public string FolderUrn { get; set; }
        public string NewExchangeName { get; set; }
        
    }
}
