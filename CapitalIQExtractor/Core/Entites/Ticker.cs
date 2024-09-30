namespace CapitalIQExtractor.Core.Entites;
using ClosedXML.Excel;

public class Ticker
{
    public string ID { get; init; }
    public string ISIN { get; init; }   
    
    public string FilePath { get; set; }
    public FormulaResult Formulas(string id)
    {
        return new FormulaResult
        {
            Price = $"CIQRANGE(\"{id}\", \"IQ_PRICEDATE\", EDATE(TODAY()-1,-24), TODAY()-1,,,\"DESC\",\"D\",\"DATA\")",
            Data = $"CIQRANGE(\"{id}\", \"IQ_CLOSEPRICE\", EDATE(TODAY()-1,-24), TODAY()-1,,,\"DESC\",\"D\",\"PRICE\")"
        };
    }

    public void MontarExcel(string folderPath)
    {
        string filePath = $"{folderPath}\\{ID}.xlsx";
        
        using (var workbook = new XLWorkbook())
        {
            var worksheet  = workbook.Worksheets.Add(ID);

            worksheet.Cell("A1").FormulaA1 = Formulas(ID).Price;
            worksheet.Cell("B1").FormulaA1 = Formulas(ID).Data;
            
            workbook.SaveAs(filePath, true); // Overwrite if exists
        }
        
        FilePath = filePath;
    }
}

public class FormulaResult
{
    public string Price { get; set; }
    public string Data { get; set; }
}