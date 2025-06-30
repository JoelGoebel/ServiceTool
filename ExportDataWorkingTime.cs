using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    public class ExportDataWorkingTime
    {
        public DateTime EinsatzDatum_Start { get; set; }
        public DateTime EinsatzDatum_Ende { get; set; }
        public string Auftragsnummer { get; set; }
        public string ServiceTechnicker { get; set; }
        public TimeSpan ArbeitsZeit_Start { get; set; }
        public TimeSpan ArbeitsZeit_Ende { get; set; }
        public TimeSpan ArbeitsZeit_Pause { get; set; }
        public TimeSpan ArbeitsZeit_Start_S2 { get; set; }
        public TimeSpan ArbeitsZeit_Ende_S2 { get; set; }
        public TimeSpan ArbeitsZeit_Pause_S2 { get; set; }
        public TimeSpan ArbeitsZeit_NormalHours { get; set; }
        public TimeSpan ArbeitsZeit_Overtime { get; set; }
        public TimeSpan ArbeitsZeit_NightWork { get; set; }
        public TimeSpan ArbeitsZeit_Gesamt { get; set; }
        public DateTime AnreiseDatum_Start { get; set; }
        public DateTime AnreiseDatum_Ende { get; set; }
        public TimeSpan Anreise_Startzeit { get; set; }
        public TimeSpan Anreise_Endezeit { get; set; }
        public TimeSpan Anreise_Pause { get; set; }
        public TimeSpan Anreise_DauerGesamt { get; set; }
        public string Anreise_KM { get; set; }
        public DateTime AbreiseDatum_Start { get; set; }
        public DateTime AbreiseDatum_Ende { get; set; }
        public TimeSpan Abreise_Startzeit { get; set; }
        public TimeSpan Abreise_Endezeit { get; set; }
        public TimeSpan Abreise_Pause { get; set; }
        public TimeSpan Abreise_DauerGesamt { get; set; }
        public string Abreise_KM { get; set; }
    }
}
