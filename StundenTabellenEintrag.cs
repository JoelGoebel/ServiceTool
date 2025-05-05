using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceTool
{
    class StundenTabellenEintrag
    {
        public string Date { get; set; }
        public string Start { get; set; }
        public string End { get; set; }
        public string Break { get; set; }
        public string StartS2 { get; set; }
        public string EndS2 { get; set; }
        public string BreakS2 { get; set; }
        public string Note { get; set; }
        public string NormalStunden { get; set; }
        public string OverTime { get; set; }
        public string Nightwork { get; set; }
        public string TotalHours { get; set; }

    }

}
