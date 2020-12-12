using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace AutomatizadorDeTransferenciaDeDados
{
    public class AbrePlanilhacs
    {
        public int NumeroDeLinhas { get; set; }
        public OutboundData[] AbrePlanilha()
        {
            int arrayLenght = 0;
            OutboundData[] excelDataInfo2 = new OutboundData[3000];
            //    var dao = new DAO();
            // int retornoDaConsulta;
            string[] vs = new string[1];
            //Declaração dos links das planilhas que o sistema deve acessar
            vs[0] = "G:/20_MAINTENANCE/02-PFMAP/Juan/Remessas KD 2020 - Edição 07.0 - PP02.3.xlsx";
            for (int i = 0; i < vs.Length; i++)
            {
                //Tenta realizar a leitura da planilha

                //Abre o arquivo no link setado 
                using (var stream = File.Open(vs[i], FileMode.Open, FileAccess.Read))
                {
                    //Abre a planilha no link setado 
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        OutboundData[] excelDataInfo = new OutboundData[reader.RowCount];
                        
                        NumeroDeLinhas = reader.RowCount;
                        int n = 0;
                        do
                        {
                            //Executa a leitura enquato existir linha e colunas a serem lidas
                            while (reader.Read())
                            {

                                if (Convert.ToString(reader.GetValue(0)).Contains("ZAT") || Convert.ToString(reader.GetValue(0)).Contains("ZAB"))
                                {
                                    var excelData = new OutboundData();
                                    excelData.BatchId = Convert.ToString(reader.GetValue(0));

                                        excelData.PopId = Convert.ToString(reader.GetValue(1));

                                        excelData.Chassis = Convert.ToString(reader.GetValue(2));

                                        excelData.CustomerOrder = Convert.ToString(reader.GetValue(3));

                                        excelData.PartPeriod = Convert.ToString(reader.GetValue(4));

                                        excelData.Type = Convert.ToString(reader.GetValue(5));

                                        excelData.Market = Convert.ToString(reader.GetValue(6));

                                        excelData.Model = Convert.ToString(reader.GetValue(27));

                                        excelData.CabType = Convert.ToString(reader.GetValue(28));

                                        excelData.CabLenght = Convert.ToString(reader.GetValue(29));

                                        excelData.RoofHeight = Convert.ToString(reader.GetValue(31));

                                        excelData.PDD = Convert.ToString(reader.GetValue(54));

                                        excelData.PlanPacking = Convert.ToString(reader.GetValue(20));

                                        excelData.PlanDelivery = Convert.ToString(reader.GetValue(20));

                                        excelData.PortDestination = "DURBAN";
                                        string IsV8OrNot = "V8";
                                        if (Convert.ToString(reader.GetValue(14)).Contains(IsV8OrNot))
                                        {
                                            excelData.InttraNumber = Convert.ToString(reader.GetValue(14));
                                        }
                                        excelDataInfo[n] = excelData;
                                        n++;
                                

                                }






                            }

                            //Retorna a leitura caso exista alguma outra aba na planilha
                        } while (reader.NextResult());
                        for (int h = 0; h < excelDataInfo.Length; h++)
                        {
                            excelDataInfo2[h] = excelDataInfo[h];
                        }
                        


                    }
                }

               
            }
            return excelDataInfo2;
        }
    }
}
