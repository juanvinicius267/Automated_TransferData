using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatizadorDeTransferenciaDeDados
{
    public class OutboundData
    {
      
        public string BatchId { get; set; }
        public string PopId { get; set; }
        public string Chassis { get; set; }
        public string CustomerOrder { get; set; }
        public string PartPeriod { get; set; }
        public string Type { get; set; }
        public string Market { get; set; }
        public string Model { get; set; }
        public string CabType { get; set; }
        public string CabLenght { get; set; }
        public string RoofHeight { get; set; }
        public string PDD { get; set; }
        public string PlanPacking { get; set; }
        public string PlanDelivery { get; set; }
        public string PortDestination { get; set; }
        public string InttraNumber { get; set; }

    }
}
