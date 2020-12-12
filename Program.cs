using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatizadorDeTransferenciaDeDados
{
    public class Program
    {
        static void Main(string[] args)
        {
            //OLEDB oLEDB = new OLEDB();
            //oLEDB.Conecc();
            AbrePlanilhacs abrePlanilha = new AbrePlanilhacs();
            OutboundData[] excelDataInfo = abrePlanilha.AbrePlanilha();
            EscreveNaPlanilhaDoOutbound escreve = new EscreveNaPlanilhaDoOutbound();
            escreve.SetDataInExcel(excelDataInfo);
        }
    }
}
