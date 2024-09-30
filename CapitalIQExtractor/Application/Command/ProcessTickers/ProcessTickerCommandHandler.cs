using CapitalIQExtractor.Core.Entites;
using CapitalIQExtractor.Core.Interfaces;
using MediatR;
using Excel = Microsoft.Office.Interop.Excel;

namespace CapitalIQExtractor.Application.Command.ProcessTickers;

public class ProcessTickerCommandHandler : IRequestHandler<ProcessTickerCommand, string>
{
    private readonly IExcelAddinService _excelAddinService;
    private const int MaxExcelInstances = 5; // Definir o tamanho máximo do pool de instâncias

    public ProcessTickerCommandHandler(IExcelAddinService excelAddinService)
    {
        _excelAddinService = excelAddinService;
    }

    public async Task<string> Handle(ProcessTickerCommand request, CancellationToken cancellationToken)
    {
        try
        {
            List<Ticker> tickers = request.Tickers;
            string folderPath = "C:\\Users\\guime\\RiderProjects\\CapitalIQExtractor\\CapitalIQExtractor\\Output\\Tickers";

            // Montar fórmula para cada ticker
            tickers.ForEach(ticker =>
            {
                ticker.MontarExcel(folderPath);
            });

            // Determinar o número necessário de instâncias de Excel, limitado ao máximo permitido
            int numInstancesNeeded = Math.Min(tickers.Count, MaxExcelInstances);

            // Criar pool de instâncias do Excel conforme a necessidade
            var excelInstances = await CreateExcelInstancePool(numInstancesNeeded);

            // Dividir tickers entre as instâncias do Excel
            var tickerGroups = DistributeTickers(tickers, excelInstances.Count);

            // Processar os tickers em paralelo usando o pool de instâncias
            var processingTasks = new List<Task>();
            for (int i = 0; i < excelInstances.Count; i++)
            {
                var excelInstance = excelInstances[i];
                var tickersToProcess = tickerGroups[i];

                processingTasks.Add(ProcessTickersAsync(excelInstance, tickersToProcess, folderPath));
            }

            // Esperar todas as tarefas de processamento terminarem
            await Task.WhenAll(processingTasks);

            // Fechar as instâncias do Excel
            CloseExcelInstances(excelInstances);

            return tickers.First().ID; // Retorna o ID do primeiro ticker (pode ser ajustado conforme necessário)

        }
        finally
        {
            // Garantir que todas as instâncias do Excel sejam fechadas
            _excelAddinService.CloseExcelProcess();
        }
    }

    // Método para criar o pool de instâncias do Excel
    private async Task<List<Excel.Application>> CreateExcelInstancePool(int instanceCount)
    {
        var excelInstances = new List<Excel.Application>();

        for (int i = 0; i < instanceCount; i++)
        {
            var excelInstance = await _excelAddinService.GetExcelInstance();
            excelInstances.Add(excelInstance);
        }

        return excelInstances;
    }

    // Método para distribuir os tickers entre as instâncias do Excel
    private List<List<Ticker>> DistributeTickers(List<Ticker> tickers, int instanceCount)
    {
        var tickerGroups = new List<List<Ticker>>();

        // Inicializar os grupos de tickers
        for (int i = 0; i < instanceCount; i++)
        {
            tickerGroups.Add(new List<Ticker>());
        }

        // Distribuir os tickers de forma balanceada entre as instâncias
        for (int i = 0; i < tickers.Count; i++)
        {
            tickerGroups[i % instanceCount].Add(tickers[i]);
        }

        return tickerGroups;
    }

    // Método assíncrono para processar tickers
    private async Task ProcessTickersAsync(Excel.Application excelInstance, List<Ticker> tickers, string folderPath)
    {
        foreach (var ticker in tickers)
        {
            await ProcessarAsync(excelInstance, ticker, folderPath);
        }
    }

    // Processar o ticker de forma assíncrona
    private async Task ProcessarAsync(Excel.Application excel, Ticker ticker, string folderPath)
    {
        await Task.Run(() =>
        {
            excel.Workbooks.Open(ticker.FilePath);
            excel.Visible = false;
        
            // Executar macro
            excel.Run("SNLXLAddin.xla!RefreshSheet");
        
            // Salvar como CSV
            excel.ActiveWorkbook.SaveAs($"{folderPath}\\{ticker.ID}.csv", Excel.XlFileFormat.xlCSV, Local: true);
        
            // Fechar workbook
            excel.ActiveWorkbook.Close();
        });
    }

    // Fechar todas as instâncias do Excel
    private void CloseExcelInstances(List<Excel.Application> excelInstances)
    {
        foreach (var excel in excelInstances)
        {
            excel.ActiveWorkbook?.Close(false);
            excel.Quit();
        }
        
        //close any excel process existente
    }
}
