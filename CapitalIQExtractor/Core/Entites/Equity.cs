using ClosedXML.Excel;

namespace CapitalIQExtractor.Core.Entites;

public class Equity(string isin)
{
    public string Isin { get; init; } = isin;
    public Ticker? Ticker { get; set; }
    public Cds? Cds { get; set; }

    public string? FilePath { get; set; }

    public void MontarExcel(string folderPath)
    {
        var filePath = $"{folderPath}\\{Isin}.xlsx";

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add(Isin);

            MontarTicker(worksheet);
            
            MontarCds(worksheet);

            workbook.SaveAs(filePath, true); // Overwrite if exists
        }

        FilePath = filePath;
    }

    private void MontarTicker(IXLWorksheet worksheet)
    {
        var tickerBuilder = new Ticker.Builder(Isin);
        
        worksheet.Cell("A1").Value = nameof(Isin);
        worksheet.Cell("A2").Value = Isin;

        worksheet.Cell("B1").Value = nameof(tickerBuilder.TickerIdFormula);
        worksheet.Cell("B2").FormulaA1 = tickerBuilder.TickerIdFormula;

        worksheet.Cell("C1").FormulaA1 = tickerBuilder.TickerPeriodFormula;
        worksheet.Cell("D1").FormulaA1 = tickerBuilder.TickerPriceFormula;
    }
    private void MontarCds(IXLWorksheet worksheet)
    {
        var cdsBuilder = new Cds.Builder(Isin);
            
        worksheet.Cell("E1").Value = nameof(cdsBuilder.CdsidFormula);
        worksheet.Cell("E2").FormulaA1 = cdsBuilder.CdsidFormula;
            
        worksheet.Cell("F1").Value = nameof(cdsBuilder.CdsNameFormula);
        worksheet.Cell("F2").FormulaA1 = cdsBuilder.CdsNameFormula;
            
        worksheet.Cell("G1").FormulaA1 = cdsBuilder.CdsPeriodFormula;
            
        worksheet.Cell("H1").Value = nameof(cdsBuilder.CdsPriceFormula);
        worksheet.Range("H2:H600").FormulaA1 = cdsBuilder.CdsPriceFormula;
    }

}