namespace CapitalIQExtractor.Core.Entites;

public class Cds
{
    public string? Cdsid { get; set; }
    
    public string? CdsName { get; set; }
    
    public List<MarketData> Cotacao { get; set; }

    public class Builder
    {
        public string CdsidFormula { get; init; }
        public string CdsNameFormula { get; init; }
        public string CdsPeriodFormula { get; init; }
        public string CdsPriceFormula { get; init; }
        private string ConstantePrefixoIsin { get; set; }
        private static string ConstanteInitialDate { get; } = "TODAY()-1";
        private static string ConstanteFinalDate { get; } = "EDATE(TODAY()-1,-24)";

        public Builder(string isin)
        {
            ConstantePrefixoIsin = $"I_{isin}";

            CdsidFormula = $"CIQ(\"{ConstantePrefixoIsin}\",\"IQ_CDS_LIST\",5)";
            CdsNameFormula = $"CIQ(CIQ(\"{ConstantePrefixoIsin}\",\"IQ_CDS_LIST\",5),\"IQ_CDS_NAME\")";
            CdsPeriodFormula =
                $"CIQRANGE(\"{ConstantePrefixoIsin}\", \"IQ_PRICEDATE\", {ConstanteFinalDate}, {ConstanteInitialDate},,,\"DESC\",\"D\",\"{nameof(CdsPriceFormula)}\")";
            CdsPriceFormula =
                $"IF(INDIRECT(ADDRESS(ROW(),COLUMN()-1))=0,\"\",CIQ(CIQ(\"{ConstantePrefixoIsin}\",\"IQ_CDS_LIST\",5), \"IQ_CDS_MID\", INDIRECT(ADDRESS(ROW(),COLUMN()-1))))";
        }
    }
}