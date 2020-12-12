using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomatizadorDeTransferenciaDeDados
{
    public class EscreveNaPlanilhaDoOutbound
    {
        public int NumeroDeLinhas { get; set; }
        public void SetDataInExcel(OutboundData[] excelDataInfo)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"G:/20_MAINTENANCE/02-PFMAP/KDFU.xls");
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();

                NumeroDeLinhas = 0;//reader.RowCount;
                int totalDeLinhas = NumeroDeLinhas + excelDataInfo.Length;
                for (int numeroLinhas = 1; numeroLinhas <= totalDeLinhas; numeroLinhas++)
                {
                    if(excelDataInfo[numeroLinhas] == null)
                    {
                        break;
                    }
                    for (int numeroColuna = 1; numeroColuna < 17; numeroColuna++)
                    {
                        // [Linha,Coluna]
                        excelWorksheet.Cells[numeroLinhas +1910, 1] = (Convert.ToString(excelDataInfo[numeroLinhas].BatchId));
                        excelWorksheet.Cells[numeroLinhas + 1910, 2] = (Convert.ToString(excelDataInfo[numeroLinhas].PopId));
                        excelWorksheet.Cells[numeroLinhas + 1910, 3] = (Convert.ToString(excelDataInfo[numeroLinhas].Chassis));
                        excelWorksheet.Cells[numeroLinhas + 1910, 4] = (Convert.ToString(excelDataInfo[numeroLinhas].CustomerOrder));
                        excelWorksheet.Cells[numeroLinhas + 1910, 5] = (Convert.ToString(excelDataInfo[numeroLinhas].PartPeriod));
                        excelWorksheet.Cells[numeroLinhas + 1910, 6] = (Convert.ToString(excelDataInfo[numeroLinhas].Type));
                        excelWorksheet.Cells[numeroLinhas + 1910, 7] = (Convert.ToString(excelDataInfo[numeroLinhas].Market));
                        excelWorksheet.Cells[numeroLinhas + 1910, 8] = (Convert.ToString(excelDataInfo[numeroLinhas].Model));
                        excelWorksheet.Cells[numeroLinhas + 1910, 9] = (Convert.ToString(excelDataInfo[numeroLinhas].CabType));
                        excelWorksheet.Cells[numeroLinhas + 1910, 10] = (Convert.ToString(excelDataInfo[numeroLinhas].CabLenght));
                        excelWorksheet.Cells[numeroLinhas + 1910, 11] = (Convert.ToString(excelDataInfo[numeroLinhas].RoofHeight));
                        excelWorksheet.Cells[numeroLinhas + 1910, 12] = (Convert.ToString(excelDataInfo[numeroLinhas].PDD));
                        excelWorksheet.Cells[numeroLinhas + 1910, 13] = (Convert.ToString(excelDataInfo[numeroLinhas].PlanPacking));
                        excelWorksheet.Cells[numeroLinhas + 1910, 14] = (Convert.ToString(excelDataInfo[numeroLinhas].PlanDelivery));
                        excelWorksheet.Cells[numeroLinhas + 1910, 16] = (Convert.ToString(excelDataInfo[numeroLinhas].PortDestination));
                        excelWorksheet.Cells[numeroLinhas + 1910, 17] = (Convert.ToString(excelDataInfo[numeroLinhas].InttraNumber));
                      
                    }
                    Console.WriteLine("Foi gravada a linha: " + numeroLinhas + " na planilha");

                }
                //excelApp.ActiveWorkbook.Save(@"G:/20_MAINTENANCE/02-PFMAP/KDFU.xls");
                excelApp.ActiveWorkbook.Save();

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
    }
}
