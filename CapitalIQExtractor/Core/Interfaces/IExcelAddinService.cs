namespace CapitalIQExtractor.Core.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;

public interface IExcelAddinService
{
    public Task<Excel.Application> GetExcelInstance();
    
    public void CloseExcelProcess();
}