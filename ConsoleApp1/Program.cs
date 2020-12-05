using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel; //importação do pacote CLOSEDXML

namespace LendoBairroeTaxas
{
    class Program
    {
        
        public static void Main(string[] args)
        {
            // Abrir arquivo excel existente
            
            var tabela = new XLWorkbook(@"C:\Users\veron\Desktop\importar_bairros\bairro-taxa.xlsx"); //coloca o caminho da pasta com o arquivo XLSM do EXCEL
            var planilha = tabela.Worksheet(1);

            Console.WriteLine("".PadRight('-'));
            Console.WriteLine("Bairros".PadRight(35) + "Taxa".PadRight(15) + "Entrega".PadRight(15));
           

            var linha = 2;
            while (true)
            {
                var bairro = planilha.Cell("a" + linha.ToString()).Value.ToString();
                var taxa = (planilha.Cell("b" + linha.ToString()).Value.ToString());
                var entrega = (planilha.Cell("c" + linha.ToString()).Value.ToString());

                if (string.IsNullOrEmpty(bairro) ||
                    string.IsNullOrEmpty(taxa) ||
                    string.IsNullOrEmpty(entrega)
                    ) break; //Se a linha estiver em branco vou brekar o andamento da leitura!

                Console.Write(bairro.PadRight(35));
                Console.WriteLine(taxa.PadRight(15) + entrega.PadRight(15));
                

                linha++;
               
            }

            Console.WriteLine("".PadRight(50, '-'));
            Console.WriteLine("Feito!");
              Console.ReadKey();
        }
    }
}
