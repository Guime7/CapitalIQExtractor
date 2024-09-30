// using CapitalIQExtractor.Core.Entites;
// using CapitalIQExtractor.Core.Interfaces;
// using MediatR;
// using Excel = Microsoft.Office.Interop.Excel;
//
// namespace CapitalIQExtractor.Application.Command.ProcessTickers;
//
// public class ProcessTickerCommandHandler : IRequestHandler<ProcessTickerCommand, string>
// {
//     private readonly IExcelAddinService _excelAddinService;
//     
//     public ProcessTickerCommandHandler(IExcelAddinService excelAddinService)
//     {
//         _excelAddinService = excelAddinService;
//     }
//     
//     public async Task<string> Handle(ProcessTickerCommand request, CancellationToken cancellationToken)
//     {
//         List<Ticker> tickers = request.Tickers;
//         
//         string folderPath = "C:\\Users\\guime\\RiderProjects\\CapitalIQExtractor\\CapitalIQExtractor\\Output\\Tickers";
//         
//         // execute montarFormula for each ticker
//         tickers.ForEach(ticker =>
//         {
//             ticker.MontarExcel(folderPath);
//         });
//
//         var excel1 = await _excelAddinService.GetExcelInstance();
//         var excel2 = await _excelAddinService.GetExcelInstance();
//         
//         var tickers1 = tickers.Take(tickers.Count / 2).ToList();
//         var tickers2 = tickers.Skip(tickers.Count / 2).ToList();
//
// // Executa o processamento em paralelo usando as duas instâncias de Excel
//         await Task.WhenAll(
//             ProcessTickersAsync(excel1, tickers1, folderPath),
//             ProcessTickersAsync(excel2, tickers2, folderPath)
//         );
//         
//         
//         
//         // Open each file
//         // tickers.ForEach(ticker =>
//         // {
//         //     excel.Workbooks.Open(ticker.FilePath);
//         //     
//         //     //run macro
//         //     excel.Run("SNLXLAddin.xla!RefreshSheet");
//         //     
//         //     //save csv split
//         //     excel.ActiveWorkbook.SaveAs($"{folderPath}\\{ticker.ID}.csv", Excel.XlFileFormat.xlCSV, Local: true);
//         //     
//         //     //close workbook
//         //     excel.ActiveWorkbook.Close();
//         // });
//         
//         //close excel sem salvar
//         excel1.Quit();
//         excel2.Quit();
//         
//         var teste = request.Tickers;
//         return teste[0].ID;
//     }
//
//     // Método assíncrono para processar tickers
//     private async Task ProcessTickersAsync(Excel.Application excelInstance, List<Ticker> tickers, string folderPath)
//     {
//         foreach (var ticker in tickers)
//         {
//             await ProcessarAsync(excelInstance, ticker, folderPath);
//         }
//     }
//     
//     // Processar o ticker de forma assíncrona
//     private async Task ProcessarAsync(Excel.Application excel, Ticker ticker, string folderPath)
//     {
//         // Coloque a lógica de processamento aqui
//         await Task.Run(() =>
//         {
//             excel.Workbooks.Open(ticker.FilePath);
//         
//             //run macro
//             excel.Run("SNLXLAddin.xla!RefreshSheet");
//         
//             //save csv split
//             excel.ActiveWorkbook.SaveAs($"{folderPath}\\{ticker.ID}.csv", Excel.XlFileFormat.xlCSV, Local: true);
//         
//             //close workbook
//             excel.ActiveWorkbook.Close();
//         });
//     }
//     
//     private void Processar(Excel.Application excel, Ticker ticker, string folderPath)
//     {
//         excel.Workbooks.Open(ticker.FilePath);
//         
//         //run macro
//         excel.Run("SNLXLAddin.xla!RefreshSheet");
//         
//         //save csv split
//         excel.ActiveWorkbook.SaveAs($"{folderPath}\\{ticker.ID}.csv", Excel.XlFileFormat.xlCSV, Local: true);
//         
//         //close workbook
//         excel.ActiveWorkbook.Close();
//     }
// }