using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaListaClientes
    {
        public string StatusResponse { get; set; }
        public int Total { get; set; }
        public string ResponseDescription { get; set; }
        public List<Cliente>? data { get; set; }

        public RespuestaListaClientes()
        {
            data = new List<Cliente>();
        }
    }
}
