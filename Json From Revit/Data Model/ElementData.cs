using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json_From_Revit.Data_Model
{
    class ElementData
    {
        public string objectId { get; set; }

        public string name { get; set; }

        public string externalId { get; set; }

        public Parameters Identity_Data { get; set; }
        public WallData Wand { get; set; }
        public RoomData Raum { get; set; }
    }
}
