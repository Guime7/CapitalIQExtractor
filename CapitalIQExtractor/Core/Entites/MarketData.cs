using System.Runtime.CompilerServices;

namespace CapitalIQExtractor.Core.Entites;

public abstract class MarketData
{
    public DateOnly Period { get; set; }
    public decimal Price { get; set; }
}