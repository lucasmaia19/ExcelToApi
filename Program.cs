using System;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelInterop
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Caminho para o arquivo Excel
            string caminhoArquivo = @"E:\Workspace\B7\turma9.xlsx";

            // Inicializa o aplicativo Excel
            Excel.Application excelApp = new Excel.Application();

            // Abre o arquivo Excel
            Excel.Workbook workbook = excelApp.Workbooks.Open(caminhoArquivo);

            // Seleciona a primeira planilha
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Índice começa em 1

            // Obtém o intervalo de células usadas na planilha
            Excel.Range range = worksheet.UsedRange;

            // Itera sobre as linhas do intervalo
            for (int row = 1; row <= range.Rows.Count; row++)
            {
                // Lê os valores das células B, C e D
                string nome = ((Excel.Range)worksheet.Cells[row, 2]).Value2?.ToString();
                string email = ((Excel.Range)worksheet.Cells[row, 3]).Value2?.ToString();
                string celular = ((Excel.Range)worksheet.Cells[row, 4]).Value2?.ToString();

                var jsonData = new
                {
                    name = nome,
                    document = "",
                    city = "",
                    email = email,
                    description = "",
                    birthData = "",
                    state = "",
                    celPhone = celular,
                    address = "",
                    typeCandidate = 0
                };

                string json = Newtonsoft.Json.JsonConvert.SerializeObject(jsonData);
                string url = "https://localhost:7289/api/Candidate/create";

                using (var httpClient = new HttpClient())
                {
                    var content = new StringContent(json, Encoding.UTF8, "application/json");
                    var response = await httpClient.PostAsync(url, content);

                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Dados enviados com sucesso para a API. Linha {row}");
                    }
                    else
                    {
                        Console.WriteLine($"Erro ao enviar dados para a API: {response.StatusCode}");
                    }
                }

            }

            // Fecha o arquivo Excel
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            // Encerra o processo Excel
            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
        }
    }
}
