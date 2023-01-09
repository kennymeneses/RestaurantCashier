using DataAccess.Respuestas;
using DataAccess.Solicitudes;
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
                var clientes = await _manejadorClientes.ObtenerClientes();

                return Ok(clientes);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        [Route("GetHistoricalClient")]
        [HttpGet]
        public async Task<IActionResult> GetHistoricalClient(string dni)
        {
            try
            {
                var clientes = await _manejadorClientes.ObtenerClientes();

                return Ok(clientes);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        //crear controller para crear cliente 

        [Route("CreateClient")]
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] ClienteInput newClient)
        {
            var responseClienteCreated = new RespuestaClienteCreado();
            
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest();
                }
                responseClienteCreated = await _manejadorClientes.CrearCliente(newClient);

                return StatusCode(201, responseClienteCreated);
            }
            catch (Exception ex)
            {
                return StatusCode(501, responseClienteCreated);
            }
        }

    }
}
