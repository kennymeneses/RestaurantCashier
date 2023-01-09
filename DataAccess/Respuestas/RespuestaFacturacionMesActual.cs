using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaFacturacionMesActual
    {
        public string? StatusResponse { get; set; }
        public string? DescriptionResponse { get; set; }
        public int Total { get; set; }
        public string LastNroTicket { get; set; }
        public List<Orden>? Data { get; set; }

        public RespuestaFacturacionMesActual()
        {
            Data = new List<Orden>();
        }
    }
}
