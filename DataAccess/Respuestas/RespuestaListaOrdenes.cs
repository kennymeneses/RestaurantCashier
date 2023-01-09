using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaListaOrdenes
    {
        public List<Orden>? data { get; set; }

        public string responseDescription { get; set; }
        public string LastNroTicket { get; set; }

        public RespuestaListaOrdenes()
        {
            data = new List<Orden>();
        }
    }
}
