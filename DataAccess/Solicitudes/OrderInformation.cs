namespace DataAccess.Solicitudes
{
    public class OrderInformation
    {
        public string NroTicket { get; set; }
        public string ClientName { get; set; }
        public string ClientDni { get; set; }
        public List<OrderSingle> ordersList {get;set;}
        public double total { get; set; }

        public OrderInformation()
        {
            ordersList = new List<OrderSingle>();
        }
    }

    public class OrderSingle
    {
        public string NameProduct { get; set; }
        public double PriceProduct { get; set; }
        public int CountProduct { get; set; }
    }
}
