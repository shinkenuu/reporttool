using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Model.CV
{
    public class Vehicle
    {
        public DateTime sampleDate { get; set; }
        public string make { get; set; }
        public string model { get; set; }
        public string version { get; set; }
        public string prodYear { get; set; }
        public string modelYear { get; set; }
        public string doors { get; set; }
        public string bodyType { get; set; }
        public double? msrpPlusOpc { get; set; }
        public double? tp { get; set; }
        public int? volume { get; set; }

        public Vehicle() { }

        public Vehicle(DateTime sampleDate, string make, string model, string version, string prodYear, string modelYear, string doors, string bodyType, float? msrpPlusOpc, float? tp, int? volume)
        {
            this.sampleDate = sampleDate;
            this.make = make;
            this.model = model;
            this.version = version;
            this.prodYear = prodYear;
            this.modelYear = modelYear;
            this.doors = doors;
            this.bodyType = bodyType;
            this.msrpPlusOpc = msrpPlusOpc;
            this.tp = tp;
            this.volume = volume;
        }

    }
}
