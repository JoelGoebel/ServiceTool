using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{

        public static class GlobalVariables
        {
            public static Dictionary<string, string> CellMapping_ServiceAnforderungen = new Dictionary<string, string>();
            public static Dictionary<string, string> CellMapping_IbnP = new Dictionary<string, string>(); 
            public static Dictionary<string, string> CellMapping_Stundenachweis = new Dictionary<string, string>();
            public static Dictionary<string, string> CellMapping_IBNP_MRS= new Dictionary<string, string>();
            public static Dictionary<string, string> CellMapping_InternerBericht = new Dictionary<string, string>();
            public static bool _FirstSiteLoadFinished { get; set; } = false;
            public static List<object> _comboBoxItemsBackup { get; set; } = new List<object> ();
            public static bool StartSiteSelected { get; set; } = false;
            public static string SelectedItemIbnP { get; set; } = "";
            public static string LastSelectedSiteIbnP { get; set; } = "";
            public static string AuftragsNR { get; set; }
            public static string Kunde { get; set; }
            public static string KundenNummer { get; set; }
            public static string Sprache_Kunde { get; set; }
            public static bool auftraginDB { get; set; }
            public static string ServiceTechnicker { get; set; }
            public static string Ansprechpartner { get; set; }
            public static string Anschrift_1 { get; set; }
            public static string Anschrift_2 { get; set; }
            public static string Anreise { get; set; }
            public static string Land { get; set; }
            public static string Material { get; set; }
            public static string Maschiene_1 { get; set; }
            public static string Maschiene_2 { get; set; }
            public static string Maschiene_3 { get; set; }
            public static string Maschiene_4 { get; set; }
            public static string Baugroeße_1 { get; set; }
            public static string Baugroeße_2 { get; set; }
            public static string Baugroeße_3 { get; set; }
            public static string Baugroeße_4 { get; set; }
            public static string MaschinenNr_1 { get; set; }
            public static string MaschinenNr_2 { get; set; }
            public static string MaschinenNr_3 { get; set; }
            public static string MaschinenNr_4 { get; set; }
            public static bool Signatur_IBN_1 { get; set; }
            public static bool Signatur_IBN_2 { get; set; }
            public static bool Signatur_IBN_3 { get; set; }
            public static bool Signatur_IBN_4 { get; set; }
            public static string[,] Zellen_Objekte { get; set; }
            public static int rowCount_ZellenObjekte { get; set; }
            public static DataTable dt { get; set; }
            public static string Pfad_AuftragsOrdner { get; set; }
            public static string Pfad_QuelleVorlagen { get; set; }
            public static string Pfad_Anhaenge { get; set; }
            public static string Pfad_Unterschriften { get; set; }
            public static bool Online_or_Offline { get; set; } //True = Online, False = Offline

            public static DateTime EndeServiceEinsatz;
            public static DateTime StartServiceEinsatz;

            public static TimeSpan FruehNacht = new TimeSpan(6, 0, 0);
            public static TimeSpan SpaetNacht = new TimeSpan(21, 0, 0);
            public static TimeSpan RegularStd = new TimeSpan(8, 30, 0);
            
            public static string LastSelectedItem_MRS { get; set; }

    }
}

