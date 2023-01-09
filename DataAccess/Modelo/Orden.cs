namespace DataAccess.Modelo
{
    public class Orden
    {
        public double OrdenId { get; set; }
        public string Fecha { get; set; }
        public double Hora { get; set; }
        public string NroTicket { get; set; }
        public string IdEmpleado { get; set; }
        public string DNI { get; set; }
        public string Nombres { get; set; }
        public string Tipo { get; set; }
        public string IdArticulo { get; set; }
        public string Descripcion { get; set; }
        public double Cantidad { get; set; }
        public double Total { get; set; }
        public double ValorEmpleado { get; set; }
        public double ValorEmpresa { get; set; }
        public string? Planilla { get; set; }
        public string? Area { get; set; }
        public string? Cargo { get; set; }
        public int Eliminado { get; set; }
    }
}
