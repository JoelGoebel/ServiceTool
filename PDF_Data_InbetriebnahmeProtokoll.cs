using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    public class PDF_Data_InbetriebnahmeProtokoll
    {
        public string Customer { get; set; }
        public string ContactPerson { get; set; }
        public string LineConfiguration { get; set; }
        public string Material { get; set; }
        public string Viscosity { get; set; }
        public string FilterType { get; set; }
        public string SerialNumber { get; set; }
        public string Preloading { get; set; }
        public string ShimpackingLR { get; set; }
        public string ShimpackingZP { get; set; }

        //Prozess Parameters
        public List<string> Pressure_P1 { get; set; } = new List<string>();
        public List<string> Pressure_P2 { get; set; } = new List<string>();
        public List<string> P { get; set; } = new List<string>();
        public List<string> MassTemperatur { get; set; } = new List<string>();
        public List<string> n_Extruder { get; set; } = new List<string>();
        public List<string> Pump { get; set; } = new List<string>();
        public List<string> Q { get; set; } = new List<string>();
        public List<string> FilterElements { get; set; } = new List<string>();
        public List<string> BackFlushLoss_1gr { get; set; } = new List<string>();
        public List<string> BackFlushLoss_10gr { get; set; } = new List<string>();
        public List<string> BackFlushLossInPercent { get; set; } = new List<string>();
        public List<string> StrokeLength { get; set; } = new List<string>();
        public List<string> BackFlushPressure { get; set; } = new List<string>();
        public List<string> DriveForce { get; set; } = new List<string>();
        public List<string> FloodingPin { get; set; } = new List<string>();

        //Screenchanger Control
        //RSF Normal
        public string WStroke_Filter_RSF_1 { get; set; }
        public string WStroke_Filter_RSF_2 { get; set; }
        public string RStroke_Filter_RSF_1 { get; set; }
        public string RStroke_Filter_RSF_2 { get; set; }
        public string CycleTime_RSF_1 { get; set; }
        public string CycleTime_RSF_2 { get; set; }
        public string WStroke2_Filter_RSF_1 { get; set; }
        public string WStroke2_Filter_RSF_2 { get; set; }
        public string RStroke2_Filter_RSF_1 { get; set; }
        public string RStroke2_Filter_RSF_2 { get; set; }
        public string PPiston_Forward_1 { get; set; }
        public string PPiston_Forward_2 { get; set; }
        public string PPiston_Backward_1 { get; set; }
        public string PPiston_Backward_2 { get; set; }
        public string PPiston_Forward_2_1 { get; set; }
        public string PPiston_Forward_2_2 { get; set; }
        public string PPiston_Backward_2_1 { get; set; }
        public string PPiston_Backward_2_2 { get; set; }
        public string NumberFilterElements_1 { get; set; }
        public string NumberFilterElements_2 { get; set; }
        public string StrokesRevolt_1 { get; set; }
        public string StrokesRevolt_2 { get; set; }
        public string PuringPiston_Forward_1 { get; set; }
        public string PuringPiston_Forward_2 { get; set; }
        public string PuringPiston_Backward_1 { get; set; }
        public string PuringPiston_Backward_2 { get; set; }

        //SFX/SFXR
        public string WStroke_Filter_SFX_1 { get; set; }
        public string WStroke_Filter_SFX_2 { get; set; }
        public string RStroke_Filter_SFX_1 { get; set; }
        public string RStroke_Filter_SFX_2 { get; set; }
        public string CycleTime_SFX_1 { get; set; }
        public string CycleTime_SFX_2 { get; set; }
        public string FloodingTime_SFX_1 { get; set; }
        public string FloodingTime_SFX_2 { get; set; }
        public string FloodingTime_Change_1 { get; set; }
        public string FloodingTime_Change_2 { get; set; }
        public string SetPressure_SFX_1 { get; set; }
        public string SetPressure_SFX_2 { get; set; }
        public string Min_Pressure_1 { get; set; }
        public string Min_Pressure_2 { get; set; }
        public string ModeOfOperation_SFX_1 { get; set; }
        public string ModeOfOperation_SFX_2 { get; set; }
        public string PreDiff_Pressure_1 { get; set; }
        public string PreDiff_Pressure_2 { get; set; }
        public string Flooding_dim_A_1 { get; set; }
        public string Flooding_dim_A_2 { get; set; }
        public string PistonCrossSection_1 { get; set; }
        public string PistonCrossSection_2 { get; set; }
        public string MeltDischarge_1 { get; set; }
        public string MeltDischarge_2 { get; set; }

        //KSF
        public string MV_A_1 { get; set; }
        public string MV_A_2 { get; set; }
        public string MV_B_1 { get; set; }
        public string MV_B_2 { get; set; }
        public string ScreenLifeTime_1 { get; set; }
        public string ScreenLifeTime_2 { get; set; }
        public string FloodingTime_KSF_1 { get; set; }
        public string FloodingTime_KSF_2 { get; set; }
        public string Pbetween_br_Plates_1 { get; set; }
        public string Pbetween_br_Plates_2 { get; set; }
        public string Set_Pressure_KSF_1 { get; set; }
        public string Set_Pressure_KSF_2 { get; set; }
        public string Min_Pressure_KSF_1 { get; set; }
        public string Min_Pressure_KSF_2 { get; set; }
        public string Mode_Of_Operation_1 { get; set; }
        public string Mode_Of_Operation_2 { get; set; }

        //VIS
        public string VIS { get; set; }
        public string dSheet { get; set; }
        public string KP { get; set; }
        public string KK { get; set; }

        //Korrekte Funktion der Steuerung
        public string Disc_Rotation { get; set; }
        public string DriveLoadMeasurement { get; set; }
        public string BackflushStrokeLength { get; set; }
        public string PhotoAttachment_Yes { get; set; }
        public string PhotoAttachment_No { get; set; }
        public string PhotoAttachment_No_Because { get; set; }

        //Temperaturprofil in Extrusionsrichtung
        public List<string> Designation_Tempprofil { get; set; } = new List<string>();
        public List<string> Temperatur_Tempprofil { get; set; } = new List<string>();
        public string Customer_Temperature_Meassurement_korrekt { get; set; }
        public string PressureCutoff { get; set; }
        public string ElectricCutoff { get; set; }
        public string MechanicCutoff { get; set; }
        public string SetTo { get; set; }
        public string SetBar { get; set; }
        public string NoCutoff { get; set; }
        public string PlaceSignature { get; set; }
    }
}
