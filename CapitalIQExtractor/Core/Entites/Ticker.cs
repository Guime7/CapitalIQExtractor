using ClosedXML.Excel;
using Microsoft.VisualBasic;

namespace CapitalIQExtractor.Core.Entites;

public class Ticker
{
    public string? TickerId { get; set; }
    public List<MarketData>? Cotacao { get; set; }

    public Ticker(List<MarketData> cotacao)
    {
        Cotacao = cotacao;
    }

    public class Builder
    {
        public string TickerIdFormula { get; set; }
        public string TickerPeriodFormula { get; init; }
        public string TickerPriceFormula { get; init; }
        private string ConstantePrefixoIsin { get; set; }
        private static string ConstanteInitialDate { get; } = "TODAY()-1";
        private static string ConstanteFinalDate { get; } = "EDATE(TODAY()-1,-24)";

        public Builder(string isin)
        {
            ConstantePrefixoIsin = $"I_{isin}";

            TickerIdFormula = $"=SPG(\"{ConstantePrefixoIsin}\", \"SP_EXCHANGE_TICKER\")";

            TickerPeriodFormula =
                $"CIQRANGE(\"{ConstantePrefixoIsin}\", \"IQ_PRICEDATE\", {ConstanteFinalDate}, {ConstanteInitialDate},,,\"DESC\",\"D\",\"{nameof(TickerPeriodFormula)}\")";
            TickerPriceFormula =
                $"CIQRANGE(\"{ConstantePrefixoIsin}\", \"IQ_CLOSEPRICE\", {ConstanteFinalDate}, {ConstanteInitialDate},,,\"DESC\",\"D\",\"{nameof(TickerPriceFormula)}\")";
        }
    }
}