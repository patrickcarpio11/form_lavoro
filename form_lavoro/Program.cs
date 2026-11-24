using System;
using System.Drawing;
using ClosedXML.Excel; //Libreria che deve essere obbligatoriamente installata con "NuGet"
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/Conversione.xlsx"; 

        // 1. CREAZIONE E SCRITTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Conversione binaria a esadecima");
            int scelta;
            worksheet.Cell(6, 6).Value= "Decimale";
            worksheet.Cell(6, 8).Value = "Esadecimale";
            worksheet.Cell(1, 4).Value = "Numero binario";


            Console.WriteLine("Scegli cosa vuoi fare (inserisci l'opzione con il rispettivo numero) : ");
            Console.WriteLine("1)Conversione da binario a decimale a esadecimale ");
            Console.WriteLine("2)Conversione da esadecimale a binario");
            
            scelta=Convert.ToInt32(Console.ReadLine()); 
            if(scelta==1)
            {
                for (int i = 1; i < 9; i++)
                {
                    do
                    {
                        Console.WriteLine("Inserisci il bit " + i);
                        worksheet.Cell(2, i).Value = Console.ReadLine();
                    }
                    while (worksheet.Cell(2, i).GetValue<string>() != "0" && worksheet.Cell(2, i).GetValue<string>() != "1");

                }
                worksheet.Cell(3, 1).FormulaA1 = "=A2 * 2^7";
                worksheet.Cell(3, 2).FormulaA1 = "=B2 * 2^6";
                worksheet.Cell(3, 3).FormulaA1 = "=C2 * 2^5";
                worksheet.Cell(3, 4).FormulaA1 = "=D2 * 2^4";
                worksheet.Cell(3, 5).FormulaA1 = "=E2 * 2^3";
                worksheet.Cell(3, 6).FormulaA1 = "=F2 * 2^2";
                worksheet.Cell(3, 7).FormulaA1 = "=G2 * 2^1";
                worksheet.Cell(3, 8).FormulaA1 = "=H2 * 2^0";

                worksheet.Cell(7, 6).FormulaA1 = "=A3+B3+C3+D3+E3+F3+G3+H3";
                int numerodecimale = worksheet.Cell("F7").GetValue<int>();

                Console.WriteLine("Il numero in decimale è : " + numerodecimale);

                string esadecimale = numerodecimale.ToString("X");

                worksheet.Cell(7, 8).Value = esadecimale;

                Console.WriteLine("Il numero in esadecimale è : " + esadecimale);
            }else if (scelta == 2)
            {
                for (int j = 0; j < 2; j++)
                {
                    Console.WriteLine("inserisci la lettera in esadecimale");

                }
            }
          
            
            workbook.SaveAs(filePath);
        }
    }
}
