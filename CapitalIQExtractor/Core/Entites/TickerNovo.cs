namespace CapitalIQExtractor.Core.Entites;
using ClosedXML.Excel;

public class Acionistas
{
    public string Nome { get; set; }
    public string Participacao { get; set; }
}

public class Cotacao
{
    public string Data { get; set; }
    public double Preco { get; set; }
}
public class Fundamentos 
{
    public string Setor { get; set; }
    public string Segmento { get; set; }
    public string SubSegmento { get; set; }
}
public class Ticker
{
    public string ID { get; init; }
    public string ISIN { get; init; }
    
    public List<Cotacao> cotacoes { get; set; }
    public Fundamentos fundamentos { get; set; }
    public List<Acionistas> acionistas { get; set; }
    
    public string FilePath { get; set; }
    public  Formulas(string id)
    {
        {
            Cotacao
            {
                Price = $"CIQRANGE(\"{id}\", \"IQ_PRICEDATE\", EDATE(TODAY()-1,-24), TODAY()-1,,,\"DESC\",\"D\",\"DATA\")",
                Data = $"CIQRANGE(\"{id}\", \"IQ_CLOSEPRICE\", EDATE(TODAY()-1,-24), TODAY()-1,,,\"DESC\",\"D\",\"PRICE\")"
            }
        };
    }
