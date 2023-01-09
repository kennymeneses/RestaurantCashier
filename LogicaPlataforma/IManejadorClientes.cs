using DataAccess.Respuestas;
using DataAccess.Solicitudes;

namespace LogicaPlataforma
{
    public interface IManejadorClientes
    {
        Task<RespuestaListaClientes> ObtenerClientes();
        Task<RespuestaClienteEspecifico> ObtenerCliente(string dni);
        Task<RespuestaClienteCreado> CrearCliente(ClienteInput input);
    }
}
