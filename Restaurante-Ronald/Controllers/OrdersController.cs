using DataAccess.Respuestas;
using DataAccess.Solicitudes;
using LogicaPlataforma;
using Microsoft.AspNetCore.Mvc;

namespace Restaurante_Ronald.Controllers
{
    [Route("api/orders")]
    [ApiController]
    public class OrdersController : ControllerBase
    {
        private readonly ILogger<WeatherForecastController> _logger;
        public IManejadorOrdenes _manejadorOrdenes;

        public OrdersController(ILogger<WeatherForecastController> logger, IManejadorOrdenes manejadorOrdenes)
        {
            _logger = logger;
            _manejadorOrdenes = manejadorOrdenes;
        }

        [Route("GetCurrentOrders")]
        [HttpGet]
        public async Task<IActionResult> GetClients()
        {
            try
            {
                var respuestaFacturacion = await _manejadorOrdenes.ObtenerFacturacionMesActual();

                return Ok(respuestaFacturacion);
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
                var clientes = await _manejadorOrdenes.ObtenerHistoricoClienteMesActual(dni);

                return Ok(clientes);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        [Route("CreateOrder")]
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] OrderInput neworder)
        {
            var responseOrder = new RespuestaOrdenCreada();
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest();
                }

                responseOrder = await _manejadorOrdenes.RegistrarNuevaOrden(neworder);
               
                return StatusCode(201, responseOrder);
            }
            catch (Exception ex)
            {
                responseOrder.ResponseDescription = ex.Message;

                return StatusCode(501, responseOrder);
            }
        }

        //CreateNewExcelAndWrite

        [Route("CreateNewDirectory")]
        [HttpPost]
        public async Task<IActionResult> CreateNewDirectory()
        {
            var responseOrder = new RespuestaOrdenCreada();
            try
            {
                responseOrder = await _manejadorOrdenes.CreateNewExcelAndWrite();

                return StatusCode(201, responseOrder);
            }
            catch (Exception ex)
            {
                responseOrder.ResponseDescription = ex.Message;

                return StatusCode(501, responseOrder);
            }
        }
    }
}
