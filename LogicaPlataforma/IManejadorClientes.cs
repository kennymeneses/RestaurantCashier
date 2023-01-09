using DataAccess.Respuestas;
using DataAccess.Solicitudes;

namespace LogicaPlataforma
{
    public interface IManejadorClientes
    {
        Task<RespuestaListaClientes> ObtenerClientes();
        Task<RespuestaClienteCreado> CrearCliente(ClienteInput input);

        Task MigrateOrdersInfoToClientsFile();
    }
}
