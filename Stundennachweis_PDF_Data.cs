using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    class Stundennachweis_PDF_Data
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
        public List<string> Report { get; set; }

        public List<StundenTabellenEintrag> Arbeitszeit { get; set; }

    }
}
