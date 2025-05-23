using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    public class Stundennachweis_PDF_Data
    {
        public string Customer { get; set; }
        public string ServiceTechnician { get; set; }
        public string Adress1 { get; set; }
        public string Adress2 { get; set; }
        public string ContactPerson { get; set; }
        public string meansofTransport { get; set; }
        public string DateArrivalStart { get; set; }
        public string DateArrivalEnd { get; set; }
        public string TimeArrivalStart { get; set; }
        public string TimeArrivalEnd { get; set; }
        public string BreakArrival { get; set; }
        public string TotalTimeArrival { get; set; }
        public string TotalKilometersArrival { get; set; }
        public string DateDepartureStart { get; set; }
        public string DateDepartureEnd { get; set; }
        public string TimeDepartureStart { get; set; }
        public string TimeDepartureEnd { get; set; }
        public string BreakDeparture { get; set; }
        public string TotalTimeDeparture { get; set; }
        public string TotalKilometersDeparture { get; set; }
        public List<string> Report { get; set; } = new List<string>();

        public List<StundenTabellenEintrag> Arbeitszeit { get; set; } = new List<StundenTabellenEintrag>();

        public TimeSpan TotalNormalHours { get; set; }
        public TimeSpan TotalOverTime { get; set; }
        public TimeSpan TotalNightWork { get; set; }
        public TimeSpan TotalHours { get; set; }

        public string SetupAufbau { get; set; }
        public string OperatingPrinciple { get; set; }
        public string BriefingControlSystem { get; set; }
        public string Sonstiges { get; set; }
        public string Troubleshooting { get; set; }
        public string OperationOfWholeEquipment { get; set; }
        public string SafetyInstructions { get; set; }
        public string Maintenance { get; set; }
        public string EvaluationProduktGood { get; set; }
        public string EvaluationProduktMid { get; set; }
        public string EvaluationProduktBad { get; set; }
        public string EvaluationSupportGood { get; set; }
        public string EvaluationSupportMid { get; set; }
        public string EvaluationSupportBad { get; set; }
        public string PlaceCustomerSignature { get; set; }
        public string Date_Technican_Signature { get; set; }
        public string Date_Customer_Signature { get; set; }
    }
}
