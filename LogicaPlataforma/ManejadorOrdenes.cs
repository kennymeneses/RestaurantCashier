using DataAccess;
using DataAccess.Modelo;
using DataAccess.Respuestas;
using DataAccess.Solicitudes;
using ESCPOS_NET;
using ESCPOS_NET.Emitters;
using ESCPOS_NET.Utilities;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace LogicaPlataforma
{
    public class ManejadorOrdenes : IManejadorOrdenes
    {
        IManejadorClientes _manejadorClientes;

        public ManejadorOrdenes(IManejadorClientes manejadorClientes)
        {
            _manejadorClientes = manejadorClientes;
        }

        public Task ObtenerMesesFacturados()
        {
            throw new NotImplementedException();
        }

        public async Task<RespuestaFacturacionMesActual> ObtenerFacturacionMesActual()
        {
            //var sheetMesActual = _accesoArchivos.ObtenertHistorialdeLiquidacionActual();

            var sheetMesActual = ObtenertHistorialdeLiquidacionActual();
            var response = new RespuestaFacturacionMesActual();

            if (sheetMesActual != null)
            {

                var listOrders = await GetListOrders(sheetMesActual);

                if (listOrders.data == null)
                {
                    response.StatusResponse = Constantes.StatusError;
                    response.DescriptionResponse = listOrders.responseDescription;
                    response.Total = 0;
                    response.Data = listOrders.data;

                    return response;
                }

                response.StatusResponse = Constantes.StatusOk;
                response.DescriptionResponse = Constantes.SucesfullResponseDescription;
                response.LastNroTicket = listOrders.LastNroTicket;
                response.Total = listOrders.data.Count;
                response.Data = listOrders.data;

                return response;
            }
            else
            {
                response.StatusResponse = Constantes.StatusError;
                response.DescriptionResponse = Constantes.FileDoesntExists;
                response.Total = 0;
                response.Data = null;

                return response;
            }
        }

        //!!!importante de implmentar por mes cambiando solo el path
        public Task<List<Orden>> ObtenerOrdenesPorMes()
        {
            throw new NotImplementedException();
        }

        public Task ObtenerOrdenPorId()
        {
            throw new NotImplementedException();
        }

        private async Task<RespuestaListaOrdenes> GetListOrders(ISheet sheet)
        {
            var response = new RespuestaListaOrdenes();
            List<Orden>? orders = new List<Orden>();
            int rowCount = sheet.LastRowNum;
            int columsCount = sheet.DefaultColumnWidth;

            try
            {
                for (int i = 1; i <= rowCount; i++)
                {
                    var orden = new Orden();
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        rowCount = i - 1;
                        break;
                    }

                    orden = GetOrderFromRow(row);

                    if (orden.Eliminado == 0)
                    {
                        orders.Add(orden);
                    }
                }

                var ordenesActivas = new List<Orden>();


                //45127726

                IRow lastRow = sheet.GetRow(rowCount);

                response.LastNroTicket = lastRow.Cells[3].StringCellValue.Trim();
                response.responseDescription = Constantes.StatusOk;
                response.data = orders;

                return response;
            }
            catch (Exception ex)
            {
                response.responseDescription = ex.Message;
                response.data = null;

                return response;
            }
        }

        private Orden GetOrderFromRow(IRow row)
        {
            var orden = new Orden();
            try
            {
                orden.OrdenId = row.Cells[0].NumericCellValue;
                orden.Fecha = row.Cells[1].StringCellValue.Trim();
                orden.Hora = row.Cells[2].StringCellValue.Trim();
                orden.NroTicket = row.Cells[3].StringCellValue.Trim();
                orden.IdEmpleado = row.Cells[4].StringCellValue.Trim();
                orden.DNI = row.Cells[5].StringCellValue.Trim();
                orden.Nombres = row.Cells[6].StringCellValue.Trim();
                orden.Tipo = row.Cells[7].StringCellValue.Trim();
                orden.IdArticulo = row.Cells[8].StringCellValue.Trim();
                orden.Descripcion = row.Cells[9].StringCellValue;//string to double?
                orden.Cantidad = row.Cells[10].NumericCellValue;
                orden.Total = row.Cells[11].NumericCellValue;
                orden.ValorEmpleado = row.Cells[12].NumericCellValue;
                orden.ValorEmpresa = row.Cells[13].NumericCellValue;
                orden.Planilla = row.Cells[14].StringCellValue.Trim();
                orden.Area = row.Cells[15].StringCellValue.Trim();
                orden.Cargo = row.Cells[16].StringCellValue.Trim();
                orden.Eliminado = (int)row.Cells[17].NumericCellValue;

                return orden;
            }
            catch (Exception ex)
            {
                orden = null;
                return orden;
            }
        }

        public async Task<RespuestaOrdenCreada> RegistrarNuevaOrden(OrderInput order)
        {
            var respuesta = new RespuestaOrdenCreada();

            string nameFile = "LIQUIDACION" + GetMonths() + ".xlsx";
            string path = Constantes.rutaLiquidaciones + GetYear() + "/" + nameFile;

            string pathClients = Constantes.rutaClientes;

            //FileMode.Create, FileAccess.Write This combination doesnt support for write new data and remove all data

            double LastId = 0;
            try
            {
                if (System.IO.File.Exists(path))
                {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet excelSheet = workbook.GetSheet("Hoja1");
                        int numRecords = excelSheet.LastRowNum;

                        IRow row = excelSheet.GetRow(numRecords);
                        LastId = row.Cells[0].NumericCellValue;

                        IRow rowIn = excelSheet.CreateRow(numRecords + 1);
                        InsertValuesInNewRow(order, rowIn, LastId);

                        var file = new FileStream(path, FileMode.Create);
                        workbook.Write(file);

                        file.Close();

                        //add nro Ticket
                        respuesta.LastNroTicket = order.NroTicket;
                        respuesta.Id = LastId + 1;
                        respuesta.ResponseStatus = Constantes.StatusOk;
                        respuesta.ResponseDescription = Constantes.OrderRegisteredSuccessfully;
                    }
                }
                else
                {
                    FileStream fs = new FileStream(path, FileMode.Create);

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Hoja1");
                    IRow headerRow = excelSheet.CreateRow(0);
                    IRow rowIn = excelSheet.CreateRow(1);
                    InsertHeaders(headerRow);
                    InsertValuesInNewRow(order, rowIn, LastId);

                    workbook.Write(fs);

                    fs.Close();

                    respuesta.LastNroTicket = order.NroTicket;
                    respuesta.Id = LastId + 1;
                    respuesta.ResponseStatus = Constantes.StatusOk;
                    respuesta.ResponseDescription = Constantes.OrderRegisteredWithNewFile;
                }

                return respuesta;
            }
            catch (Exception ex)
            {
                respuesta.Id = 0;
                respuesta.ResponseStatus = "Error";
                respuesta.ResponseDescription = ex.Message;

                return respuesta;
            }
        }

        public async Task<RespuestaHistoricoActualCliente> ObtenerHistoricoClienteMesActual(string dni)
        {
            var response = new RespuestaHistoricoActualCliente();

            try
            {
                var sheetMesActual = ObtenertHistorialdeLiquidacionActual();


                if (sheetMesActual != null)
                {

                    var listOrders = await GetListOrders(sheetMesActual);

                    if (listOrders.data == null)
                    {
                        response.ResponseStatus = Constantes.StatusError;
                        response.ResponseDescription = listOrders.responseDescription;
                        response.Total = 0;
                        response.data = listOrders.data;

                        return response;
                    }

                    response.ResponseStatus = Constantes.StatusOk;
                    response.ResponseDescription = Constantes.SucesfullResponseDescription;

                    var ListordersByDni = listOrders.data.Where(x => x.DNI == dni).ToList();

                    response.Total = ListordersByDni.Count;
                    response.data = ListordersByDni;

                    return response;
                }
                else
                {
                    response.ResponseStatus = Constantes.StatusError;
                    response.ResponseDescription = Constantes.FileDoesntExists;
                    response.Total = 0;
                    response.data = null;

                    return response;
                }


                return response;
            }
            catch (Exception ex)
            {

                throw;
            }

        }

        private void InsertNewClient(OrderInput order, IRow rowIn, double LastId)
        {
            ICell cellIn1 = rowIn.CreateCell(0);
            cellIn1.SetCellValue(LastId);

            ICell cellIn2 = rowIn.CreateCell(1);
            cellIn2.SetCellValue(order.DNI);

            ICell cellIn3 = rowIn.CreateCell(2);
            cellIn3.SetCellValue(order.Nombres);

            ICell cellIn4 = rowIn.CreateCell(3);
            cellIn4.SetCellValue(order.Planilla);

            ICell cellIn5 = rowIn.CreateCell(4);
            cellIn5.SetCellValue(order.Area);

            ICell cellIn6 = rowIn.CreateCell(5);
            cellIn6.SetCellValue(order.Cargo);

            ICell cellIn7 = rowIn.CreateCell(6);
            cellIn7.SetCellValue(0);
        }

        private void InsertValuesInNewRow(OrderInput order, IRow rowIn, double LastId)
        {
            ICell cellIn1 = rowIn.CreateCell(0);
            cellIn1.SetCellValue(LastId + 1);

            ICell cellIn2 = rowIn.CreateCell(1);
            cellIn2.SetCellValue(order.Fecha);

            ICell cellIn3 = rowIn.CreateCell(2);
            cellIn3.SetCellValue(order.Hora);

            ICell cellIn4 = rowIn.CreateCell(3);
            cellIn4.SetCellValue(order.NroTicket);

            ICell cellIn5 = rowIn.CreateCell(4);
            cellIn5.SetCellValue(order.IdEmpleado);

            ICell cellIn6 = rowIn.CreateCell(5);
            cellIn6.SetCellValue(order.DNI);

            ICell cellIn7 = rowIn.CreateCell(6);
            cellIn7.SetCellValue(order.Nombres);

            ICell cellIn8 = rowIn.CreateCell(7);
            cellIn8.SetCellValue(order.Tipo);

            ICell cellIn9 = rowIn.CreateCell(8);
            cellIn9.SetCellValue(order.IdArticulo);

            ICell cellIn10 = rowIn.CreateCell(9);
            cellIn10.SetCellValue(order.Descripcion);

            ICell cellIn11 = rowIn.CreateCell(10);
            cellIn11.SetCellValue(order.Cantidad);

            ICell cellIn12 = rowIn.CreateCell(11);
            cellIn12.SetCellValue(order.Total);

            ICell cellIn13 = rowIn.CreateCell(12);
            cellIn13.SetCellValue(order.ValorEmpleado);

            ICell cellIn14 = rowIn.CreateCell(13);
            cellIn14.SetCellValue(order.ValorEmpresa);

            ICell cellIn15 = rowIn.CreateCell(14);
            cellIn15.SetCellValue(order.Planilla);

            ICell cellIn16 = rowIn.CreateCell(15);
            cellIn16.SetCellValue(order.Area);

            ICell cellIn17 = rowIn.CreateCell(16);
            cellIn17.SetCellValue(order.Cargo);

            ICell cellIn18 = rowIn.CreateCell(17);
            cellIn18.SetCellValue(0);
        }

        private void InsertHeaders(IRow rowIn)
        {
            ICell cellIn1 = rowIn.CreateCell(0);
            cellIn1.SetCellValue("IdRegistro");

            ICell cellIn2 = rowIn.CreateCell(1);
            cellIn2.SetCellValue("Fecha");

            ICell cellIn3 = rowIn.CreateCell(2);
            cellIn3.SetCellValue("Hora");

            ICell cellIn4 = rowIn.CreateCell(3);
            cellIn4.SetCellValue("NroTicket");

            ICell cellIn5 = rowIn.CreateCell(4);
            cellIn5.SetCellValue("IdEmpleado");

            ICell cellIn6 = rowIn.CreateCell(5);
            cellIn6.SetCellValue("DNI");

            ICell cellIn7 = rowIn.CreateCell(6);
            cellIn7.SetCellValue("Nombres");

            ICell cellIn8 = rowIn.CreateCell(7);
            cellIn8.SetCellValue("Tipo");

            ICell cellIn9 = rowIn.CreateCell(8);
            cellIn9.SetCellValue("IdArticulo");

            ICell cellIn10 = rowIn.CreateCell(9);
            cellIn10.SetCellValue("Descripcion");

            ICell cellIn11 = rowIn.CreateCell(10);
            cellIn11.SetCellValue("Cantidad");

            ICell cellIn12 = rowIn.CreateCell(11);
            cellIn12.SetCellValue("Total");

            ICell cellIn13 = rowIn.CreateCell(12);
            cellIn13.SetCellValue("Valor empleado");

            ICell cellIn14 = rowIn.CreateCell(13);
            cellIn14.SetCellValue("Valor empresa");

            ICell cellIn15 = rowIn.CreateCell(14);
            cellIn15.SetCellValue("Planilla");

            ICell cellIn16 = rowIn.CreateCell(15);
            cellIn16.SetCellValue("Area");

            ICell cellIn17 = rowIn.CreateCell(16);
            cellIn17.SetCellValue("Cargo");

            ICell cellIn18 = rowIn.CreateCell(17);
            cellIn17.SetCellValue("Eliminado");
        }

        public async Task<RespuestaOrdenCreada> CreateNewExcelAndWrite()
        {
            var respuesta = new RespuestaOrdenCreada();
            string guid = Guid.NewGuid().ToString();

            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString();
            string path = Constantes.rutaPruebas;
            var tempfileName = "Excel" + guid + ".xlsx";

            try
            {
                if (!Directory.Exists(path + "nuevacarpeta"))
                {
                    Directory.CreateDirectory(path + "nuevacarpeta");

                    FileStream fs = new FileStream(Path.Combine(path + "nuevacarpeta/", tempfileName), FileMode.Create);

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Hoja1");
                    IRow rowIn = excelSheet.CreateRow(0);
                    ICell cellIn1 = rowIn.CreateCell(0);
                    cellIn1.SetCellValue(year);

                    workbook.Write(fs);

                    fs.Close();

                    //respuesta.LastNroTicket = 
                    respuesta.ResponseDescription = Constantes.OrderRegisteredSuccessfully;
                    respuesta.ResponseStatus = Constantes.StatusOk;
                    respuesta.Id = 1;

                    return respuesta;
                }

                return respuesta;
            }
            catch (Exception ex)
            {
                respuesta.ResponseStatus = Constantes.StatusError;
                respuesta.ResponseDescription = ex.Message;
                return respuesta;
            }
        }

        public ISheet ObtenertHistorialdeLiquidacionActual()
        {
            string nameFile = "LIQUIDACION" + GetMonths() + ".xlsx";
            string path = Constantes.rutaLiquidaciones + GetYear() + "/" + nameFile;

            ISheet? worksheet = null;
            try
            {
                if (System.IO.File.Exists(path))
                {
                    IWorkbook? workbook = null;  //IWorkbook determina si es xls o xlsx              

                    using (FileStream FS = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);          //Abre tanto XLS como XLSX
                        worksheet = workbook.GetSheetAt(0);    //Obtener Hoja por indice                        
                    }

                    return worksheet;
                }
                return worksheet;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private string GetMonths()
        {
            int day = DateTime.Now.Day;
            int december = 12;
            int month = DateTime.Now.Month;

            string response = string.Empty;

            if (month == 1 && day <= 14)
            {
                response = december.ToString() + "-" + month.ToString();
            }

            else if (day > 14)
            {
                response = month.ToString() + "-" + (month + 1).ToString();
            }

            else if (day <= 14)
            {
                response = (month - 1).ToString() + "-" + (month).ToString();
            }

            return response;
        }

        private string GetYear()
        {
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;


            if (day >= 15 && month == 12)
            {
                year = year + 1;
            }
            return year.ToString();
        }

        private async Task<bool> ClientFromOrderExists(OrderInput order)
        {
            var respuestaClientes = await _manejadorClientes.ObtenerClientes();
            var listaClientes = respuestaClientes.data;

            bool existeCliente = listaClientes.Exists(x => x.Dni == order.DNI);

            return existeCliente;
        }

        public async Task<RespuestaListaMenus> ObtenerMenusList()
        {
            var response = new RespuestaListaMenus();

            try
            {
                var hojamenus = ObtenertHojaMenus();
                if (hojamenus == null)
                {
                    response.menus = null;
                    response.Total = 0;
                    response.ResponseStatus = Constantes.StatusError;


                    return response;
                }

                var listaMenus = await GetListMenus(hojamenus);

                response.menus = listaMenus;
                response.Total = listaMenus.Count;
                response.ResponseStatus = Constantes.StatusOk;

                return response;

            }
            catch (Exception ex)
            {
                response.menus = null;
                response.Total = 0;
                response.ResponseStatus = ex.Message;

                return response;
            }

        }

        public ISheet ObtenertHojaMenus()
        {
            string nameFile = "LIQUIDACION" + GetMonths() + ".xlsx";
            string path = Constantes.rutaMenus;

            ISheet? worksheet = null;
            try
            {
                if (System.IO.File.Exists(path))
                {
                    IWorkbook? workbook = null;  //IWorkbook determina si es xls o xlsx              

                    using (FileStream FS = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);          //Abre tanto XLS como XLSX
                        worksheet = workbook.GetSheetAt(0);    //Obtener Hoja por indice                        
                    }

                    return worksheet;
                }
                return worksheet;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private async Task<List<Menu>> GetListMenus(ISheet sheet)
        {
            var response = new RespuestaListaOrdenes();
            List<Menu>? menus = new List<Menu>();
            int rowCount = sheet.LastRowNum;
            int columsCount = sheet.DefaultColumnWidth;

            try
            {
                for (int i = 0; i <= rowCount; i++)
                {
                    var menu = new Menu();
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        rowCount = i - 1;
                        break;
                    }

                    menu = await GetMenuFromRow(row);

                    menus.Add(menu);
                }

                return menus;
            }
            catch (Exception ex)
            {
                return menus;
            }
        }

        private async Task<Menu> GetMenuFromRow(IRow row)
        {
            var menu = new Menu();
            try
            {
                menu.IdMenu = (int)row.Cells[0].NumericCellValue;
                menu.NombreMenu = row.Cells[1].StringCellValue;
                menu.precio = row.Cells[2].NumericCellValue;

                return menu;

            }
            catch (Exception ex)
            {
                menu = null;
                return menu;
            }
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            var printFont = new Font("Arial", 8);
            float linesPerPage = 0;
            int count = 0;
            float yPos = 0;
            //float leftMargin = ev.MarginBounds.Left;
            float leftMargin = 0;
            //float topMargin = ev.MarginBounds.Top;
            float topMargin = 0;
            string line = null;


            var orderInformation = new OrderInformation();
            var stream = GenerateStreamFromObject(orderInformation);

            StreamReader streamToPrint = new StreamReader(stream);

            linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);

            while (count < linesPerPage && ((line = streamToPrint.ReadLine()) != null))
            {
                yPos = topMargin + (count *
                   printFont.GetHeight(ev.Graphics));
                ev.Graphics.DrawString(line, printFont, Brushes.Black,
                   leftMargin, yPos, new StringFormat());
                count++;
            }
        }


        public Stream GenerateStreamFromObject(OrderInformation information)
        {
            StringBuilder stringBuilder = new StringBuilder("           FUEGO, SAZON Y SABOR\n");
            stringBuilder.Append("          RUC: 10329022395 \n");
            stringBuilder.Append("--------------------------------------- \n");
            stringBuilder.Append("Ticket Nro. CJ01-0099436 \n");
            stringBuilder.AppendFormat("Fecha: {0}          Hora: {1} \n", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString());
            stringBuilder.Append("----------------------------------------------- \n");
            stringBuilder.Append("Item              Cant    Precio      Total \n");
            foreach (var order in information.ordersList)
            {
                double totalProduct = order.PriceProduct * order.CountProduct;

                stringBuilder.AppendFormat("{0}     {1}     {2}     {3} \n", order.NameProduct.ToString(),
                                                            order.CountProduct.ToString(), order.PriceProduct.ToString(), totalProduct.ToString());
            }
            //stringBuilder.Append("MENU PARA LLEVAR   1      S/. 8,00    S/. 8,00 \n");
            stringBuilder.Append("----------------------------------------------- \n");
            stringBuilder.AppendFormat("Total Venta                           S/. {0}0 \n", information.total.ToString());
            stringBuilder.Append(" \n");
            stringBuilder.Append("Condicion de pago: CREDITO \n");
            stringBuilder.AppendFormat("DNI              : {0} \n", information.ClientDni.ToString());
            stringBuilder.AppendFormat("Nombres          : {0} \n", information.ClientName.ToString()); 

            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(stringBuilder);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        public async Task<ResponsePrinter> PrintOrder(OrderInformation information)
        {
            var response = new ResponsePrinter();

            try
            {
                var fileToPrinter = new PrintDocument();
                fileToPrinter.DocumentName = "documento";
                //fileToPrinter.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
                fileToPrinter.PrintPage += (sender, e) => {

                    StringBuilder stringBuilder = new StringBuilder("           FUEGO, SAZON Y SABOR\n");
                    stringBuilder.Append("          RUC: 10729022395 \n");
                    stringBuilder.Append("--------------------------------------- \n");
                    stringBuilder.AppendFormat("Ticket Nro. {0} \n", information.NroTicket);
                    stringBuilder.AppendFormat("Fecha: {0}          Hora: {1} \n", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString());
                    stringBuilder.Append("----------------------------------------------- \n");
                    stringBuilder.Append("Item              Cant    Precio      Total \n");
                    foreach (var order in information.ordersList)
                    {
                        double totalProduct = order.PriceProduct * order.CountProduct;

                        stringBuilder.AppendFormat("{0}     {1}     {2}     {3} \n", order.NameProduct.ToString(),
                                                                    order.CountProduct.ToString(), order.PriceProduct.ToString(), totalProduct.ToString());
                    }
                    //stringBuilder.Append("MENU PARA LLEVAR   1      S/. 8,00    S/. 8,00 \n");
                    stringBuilder.Append("----------------------------------------------- \n");
                    stringBuilder.AppendFormat("Total Venta                           S/. {0} \n", information.total.ToString());
                    stringBuilder.Append(" \n");
                    stringBuilder.Append("Condicion de pago: CREDITO \n");
                    stringBuilder.AppendFormat("DNI              : {0} \n", information.ClientDni.ToString());
                    stringBuilder.AppendFormat("Nombres          : {0} \n", information.ClientName.ToString());

                    MemoryStream stream = new MemoryStream();
                    StreamWriter writer = new StreamWriter(stream);
                    writer.Write(stringBuilder);
                    writer.Flush();
                    stream.Position = 0;

                    var printFont = new Font("Arial", 8);
                    float linesPerPage = 0;
                    int count = 0;
                    float yPos = 0;
                    //float leftMargin = ev.MarginBounds.Left;
                    float leftMargin = 0;
                    //float topMargin = ev.MarginBounds.Top;
                    float topMargin = 0;
                    string line = null;

                    linesPerPage = e.MarginBounds.Height / printFont.GetHeight(e.Graphics);
                    StreamReader streamToPrint = new StreamReader(stream);


                    while (count < linesPerPage && ((line = streamToPrint.ReadLine()) != null))
                    {
                        yPos = topMargin + (count *
                           printFont.GetHeight(e.Graphics));
                        e.Graphics.DrawString(line, printFont, Brushes.Black,
                           leftMargin, yPos, new StringFormat());
                        count++;
                    }
                };

                fileToPrinter.Print();

                response.status = Constantes.StatusOk;
                return response;
            }
            catch (Exception ex)
            {
                response.status = ex.Message;

                return response;             
            }            
        }
    }
}
