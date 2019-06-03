using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json_From_Revit.Data_Model
{
    class ProjectData
    {
        public string VersionName { get; set; }
        public string Architecture_Document { get; set; }
        public string Elements_Document { get; set; }

        public DocumentInformation Document_Information { get; set; }

        public List<ElementData> elements { get; set; }
    }
}
