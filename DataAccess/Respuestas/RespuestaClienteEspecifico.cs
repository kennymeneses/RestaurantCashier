using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaClienteEspecifico
    {
        public Cliente cliente { get; set; }
        public string ResponseStatus { get; set; }
        public string ResponseDescription { get; set; }
    }
}
