using CapitalIQExtractor.Core.Entites;
using MediatR;

namespace CapitalIQExtractor.Application.Command.ProcessTickers;

public record ProcessTickerCommand() : IRequest<string>
{
    public List<Ticker> Tickers { get; init; }
}

