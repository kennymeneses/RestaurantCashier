using DataAccess.Modelo;
using DataAccess.Respuestas;
using DataAccess.Solicitudes;

namespace LogicaPlataforma
{
    public interface IManejadorOrdenes
    {
        Task ObtenerMesesFacturados();
        Task<RespuestaFacturacionMesActual> ObtenerFacturacionMesActual();
        Task<List<Orden>> ObtenerOrdenesPorMes();
        Task ObtenerOrdenPorId();
        Task<RespuestaOrdenCreada> RegistrarNuevaOrden(OrderInput order);

        Task<RespuestaOrdenCreada> CreateNewExcelAndWrite();
        Task<RespuestaHistoricoActualCliente> ObtenerHistoricoClienteMesActual(string dni);
    }
}
