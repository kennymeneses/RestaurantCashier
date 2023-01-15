using DataAccess.Modelo;

namespace DataAccess.Respuestas
{
    public class RespuestaListaMenus
    {
        public int Total { get; set; }
        public string ResponseStatus { get; set; }
        public List<Menu> menus { get; set; }

        public RespuestaListaMenus()
        {
            menus = new List<Menu>();
        }
    }
}
