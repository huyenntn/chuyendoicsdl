using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AVDApplication
{
    class OutFormatBE
    {
        //STT
        public Int64 Index { get; set; }

        //GP Number
        public String GPNo { get; set; }

        //GP Type
        public String GPType { get; set; }

        //Ref no
        public String RefNo { get; set; }

        //F bias
        public String FBias { get; set; }

        //frequency
        public Double Frequency { get; set; }

        //priority band
        public String PrioBand { get; set; }
        
        //bandwidth
        public String BandWidth { get; set; }

        //number of chanel
        public String NoChanel { get; set; }

        //Customer Name
        public String CustomerName { get; set; }

        //Call
        public String Call { get; set; }

        //longitude
        public Double Longitude { get; set; }

        //latitude
        public Double Latitude { get; set; }

        //Machine Name
        public String MachineName { get; set; }
    }
}
