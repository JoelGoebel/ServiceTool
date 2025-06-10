using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    public class PDF_Data_IbnP_MRS
    {
        public string Customer { get; set; }
        public string  ContactPerson { get; set; }
        public string LineConfiguration { get; set; }
        public string Material { get; set; }
        public string ExtruderType { get; set; }
        public string SerialNumber { get; set; }
        public string FinalProduct { get; set; }
        public string Shape { get; set; }

        //Tabelle Process Parameters
        public List<string> Time { get; set; } = new List<string>();
        public List<string> Control { get; set; } = new List<string>();
        public List<string> Pump { get; set; } = new List<string>();
        public List<string> Load { get; set; } = new List<string>();
        public List<string> Extruderspeed_Soll { get; set; } = new List<string>();
        public List<string> Extruderspeed_Min { get; set; } = new List<string>();
        public List<string> Load_2 { get; set; } = new List<string>();
        public List<string> Vacuum { get; set; } = new List<string>();
        public List<string> Viscosimeter_Viscosity { get; set; } = new List<string>();
        public List<string> Viscosimeter_Shearrate { get; set; } = new List<string>();
        public List<string> MP1 { get; set; } = new List<string>();
        public List<string> MP2 { get; set; } = new List<string>();
        public List<string> MP3 { get; set; } = new List<string>();
        public List<string> MP4 { get; set; } = new List<string>();
        public List<string> MP5 { get; set; } = new List<string>();
        public List<string> Filter_P { get; set; } = new List<string>();
        public List<string> FilterFineness { get; set; } = new List<string>();
        public List<string> Screwcooling_Actual { get; set; } = new List<string>();
        public List<string> Feedzone_Cooling { get; set; } = new List<string>();
        public List<string> TM_Filter { get; set; } = new List<string>();
        public List<string> TM_Visco { get; set; } = new List<string>();
        public List<string> Throughput { get; set; } = new List<string>();


        //Tabelle Heating / Cooling
        public List<string> HZ1 { get; set; } = new List<string>();
        public List<string> HZ2 { get; set; } = new List<string>();
        public List<string> HZ3 { get; set; } = new List<string>();
        public List<string> HZ4 { get; set; } = new List<string>();
        public List<string> HZ5 { get; set; } = new List<string>();
        public List<string> HZ6 { get; set; } = new List<string>();
        public List<string> HZ7 { get; set; } = new List<string>();
        public List<string> HZ8 { get; set; } = new List<string>();
        public List<string> HZ9 { get; set; } = new List<string>();
        public List<string> HZ10 { get; set; } = new List<string>();
        public List<string> HZ11 { get; set; } = new List<string>();
        public List<string> HZ12 { get; set; } = new List<string>();
        public List<string> HZ13 { get; set; } = new List<string>();
        public List<string> HZ14 { get; set; } = new List<string>();
        public List<string> HZ15 { get; set; } = new List<string>();
        public List<string> HZ16 { get; set; } = new List<string>();
        public List<string> HZ17 { get; set; } = new List<string>();
        public List<string> HZ18 { get; set; } = new List<string>();
        public List<string> HZ19 { get; set; } = new List<string>();
        public List<string> HZ20 { get; set; } = new List<string>();
        public List<string> HZ21 { get; set; } = new List<string>();
        public List<string> HZ22 { get; set; } = new List<string>();
        public List<string> HZ23 { get; set; } = new List<string>();
        public List<string> HZ24 { get; set; } = new List<string>();
        public List<string> HZ25 { get; set; } = new List<string>();
        public List<string> HZ26 { get; set; } = new List<string>();

        public string Cooling_Feeding_Zone { get; set; }
        public string Screwcooling { get; set; }
        public string ChillerVacuum { get; set; }
        public string Remarks { get; set; }

        //Controll MRS

        public string Extruder { get; set; }
        public string Viscosimeter { get; set; }
        public string Vacuum_Control { get; set; }
        public string OtherFixParameterSettings { get; set; }

        public string Place_Signature { get; set; }
        public string Date_Signature { get; set; }
    }
}
