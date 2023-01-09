using LogicaPlataforma;
using Microsoft.AspNetCore.Mvc;

namespace Restaurante_Ronald.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ClientsController : ControllerBase
    {
        private readonly ILogger<WeatherForecastController> _logger;
        public IManejadorClientes _manejadorClientes;

        public ClientsController(ILogger<WeatherForecastController> logger, IManejadorClientes manejadorClientes)
        {
            _logger = logger;
            _manejadorClientes = manejadorClientes;
        }


        [Route("GetClients")]
        [HttpGet]
        public async Task<IActionResult> GetClients()
        {
            try
            {
                var respuestaFacturacion = await _manejadorClientes.ObtenerClientes();

                return Ok(respuestaFacturacion);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
