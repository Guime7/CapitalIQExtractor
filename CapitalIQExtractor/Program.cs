using MediatR;
using CapitalIQExtractor.Application.Command.ProcessTickers;
using CapitalIQExtractor.Core.Interfaces;
using CapitalIQExtractor.Infra.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

//Add mediator
builder.Services.AddMediatR(options =>
{
    options.RegisterServicesFromAssemblies(typeof(Program).Assembly);
});

//dependency injection
builder.Services.AddScoped<IExcelAddinService, ExcelAddinService>();


var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

#region Endpoint

app.MapPost("/process-tickers", async (ProcessTickerCommand command, ISender sender) =>
{
    return await sender.Send(command);
});


#endregion

app.Run();
