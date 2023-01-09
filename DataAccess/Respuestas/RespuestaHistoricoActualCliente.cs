using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaHistoricoActualCliente
    {
        public string ResponseStatus { get; set; }
        public string ResponseDescription { get; set; }
        public int Total { get; set; }
        public List<Orden> data { get; set; }



        public RespuestaHistoricoActualCliente()
        {
            data= new List<Orden>();
        }
    }
}
