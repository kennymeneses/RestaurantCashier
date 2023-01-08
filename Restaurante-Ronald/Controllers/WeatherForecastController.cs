using DataAccess;
using DataAccess.Modelo;
using LogicaPlataforma;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;

namespace Restaurante_Ronald.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;
        public IManejadorOrdenes _manejadorOrdenes;

        public WeatherForecastController(ILogger<WeatherForecastController> logger, IManejadorOrdenes manejadorOrdenes)
        {
            _logger = logger;
            _manejadorOrdenes = manejadorOrdenes;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }
    }
}