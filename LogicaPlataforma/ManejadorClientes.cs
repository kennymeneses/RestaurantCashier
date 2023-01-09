using DataAccess;
using DataAccess.Modelo;
using DataAccess.Respuestas;
using DataAccess.Solicitudes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace LogicaPlataforma
{
    public class ManejadorClientes : IManejadorClientes
    {
        public ManejadorClientes()
        {

        }

        public async Task<RespuestaClienteCreado> CrearCliente(ClienteInput input)
        {
            string path = Constantes.rutaClientes;
            var respuesta = new RespuestaClienteCreado();
            int LastId = 0;

            if (System.IO.File.Exists(path))
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(fs);
                    ISheet excelSheet = workbook.GetSheet("Hoja1");
                    int numRecords = excelSheet.LastRowNum;

                    IRow row = excelSheet.GetRow(numRecords);
                    LastId = (int)row.Cells[0].NumericCellValue;

                    IRow rowIn = excelSheet.CreateRow(numRecords + 1);
                    InsertCient(input, rowIn, LastId);

                    var file = new FileStream(path, FileMode.Create);
                    workbook.Write(file);

                    file.Close();

                    respuesta.Id = LastId + 1;
                    respuesta.ResponseStatus = Constantes.StatusOk;
                    respuesta.ResponseDescription = Constantes.ClienteCreatedSuccessfully;
                }                 
            }
            return respuesta;
        }

        public async Task<RespuestaListaClientes> ObtenerClientes()
        {
            var response = new RespuestaListaClientes();
            try
            {
                var hojaClientes = ObtenertHojaClientes();

                var clientList = await GetClienteList(hojaClientes);
                
                if(clientList.data == null)
                {
                    response.data = null;
                    response.Total = 0;
                    response.ResponseDescription = Constantes.ListClientEmpty;
                    response.StatusResponse = clientList.StatusResponse;

                    return response;
                }

                response.data = clientList.data;
                response.Total = clientList.Total;
                response.ResponseDescription = clientList.ResponseDescription;
                response.StatusResponse = clientList.StatusResponse;

                return response;
            }
            catch (Exception ex)
            {
                response.data = null;
                response.Total = 0;
                response.ResponseDescription = ex.Message;
                response.StatusResponse = Constantes.StatusError;

                return response;
            }
        }

        private ISheet ObtenertHojaClientes()
        {
            string path = Constantes.rutaClientes;

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

        private async Task<RespuestaListaClientes> GetClienteList(ISheet sheet)
        {
            var response = new RespuestaListaClientes();
            List<Cliente>? clientes = new List<Cliente>();
            int rowCount = sheet.LastRowNum;
            int columsCount = sheet.DefaultColumnWidth;

            try
            {
                for (int i = 1; i <= rowCount; i++)
                {
                    var cliente = new Cliente();
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        rowCount = i - 1;
                        break;
                    }

                    cliente = GetClienteFromRow(row);
                    clientes.Add(cliente);
                }

                response.StatusResponse = Constantes.StatusOk;
                response.ResponseDescription = Constantes.GetListClientSuccessfully;
                response.Total = clientes.Count;
                response.data = clientes;

                return response;
            }
            catch (Exception ex)
            {
                response.ResponseDescription = ex.Message;
                response.data = null;

                return response;
            }
        }

        private Cliente GetClienteFromRow(IRow row)
        {
            var cliente = new Cliente();
            try
            {
                cliente.IdEmpleado = row.Cells[0].NumericCellValue;
                cliente.Dni = row.Cells[1].StringCellValue.Trim();
                cliente.Nombres = row.Cells[2].StringCellValue.Trim();
                cliente.Planilla = row.Cells[3].StringCellValue.Trim();
                cliente.Area = row.Cells[4].StringCellValue.Trim();
                cliente.Cargo = row.Cells[5].StringCellValue.Trim();
                cliente.Eliminado = row.Cells[6].NumericCellValue;
                return cliente;
            }
            catch (Exception ex)
            {
                cliente = null;
                return cliente;
            }
        }


        public async Task MigrateOrdersInfoToClientsFile()
        {
            var sheetMesActual = ObtenertHistorialdeLiquidacionActual();
            var orders = await GetListOrders(sheetMesActual);
            MigrateToClients(orders);
        }

        private async Task<RespuestaListaOrdenes> GetListOrders(ISheet sheet)
        {
            var response = new RespuestaListaOrdenes();
            List<Orden>? orders = new List<Orden>();
            int rowCount = sheet.LastRowNum;
            int columsCount = sheet.DefaultColumnWidth;

            try
            {
                for (int i = 1; i < rowCount; i++)
                {
                    var orden = new Orden();
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        rowCount = i - 1;
                        break;
                    }

                    orden = GetOrderFromRow(row);

                    orders.Add(orden);
                }
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
                orden.Fecha = row.Cells[1].DateCellValue.ToString();
                orden.Hora = row.Cells[2].NumericCellValue;
                orden.NroTicket = row.Cells[3].StringCellValue.Trim();
                orden.IdEmpleado = row.Cells[4].NumericCellValue.ToString();
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

                return orden;
            }
            catch (Exception ex)
            {
                orden = null;
                return orden;
            }
        }

        public ISheet ObtenertHistorialdeLiquidacionActual()
        {
            string nameFile = "LIQUIDACION12-1.xlsx";
            string path = Constantes.rutaLiquidaciones + "2023/" + nameFile;

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

        public void MigrateToClients(RespuestaListaOrdenes listaOrdenes)
        {
            try
            {
                string ruta = "./../../excels/Clientes/prueba/Clientes.xlsx";
                FileStream fs = new FileStream(ruta, FileMode.Create);

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Hoja1");

                    IRow headerRow = excelSheet.CreateRow(0);

                    ICell cellIn1 = headerRow.CreateCell(0);
                    cellIn1.SetCellValue("IdEmpleado");

                    ICell cellIn2 = headerRow.CreateCell(1);
                    cellIn2.SetCellValue("Dni");

                    ICell cellIn3 = headerRow.CreateCell(2);
                    cellIn3.SetCellValue("Nombres");

                    ICell cellIn4 = headerRow.CreateCell(3);
                    cellIn4.SetCellValue("Planilla");

                    ICell cellIn5 = headerRow.CreateCell(4);
                    cellIn5.SetCellValue("Area");

                    ICell cellIn6 = headerRow.CreateCell(5);
                    cellIn6.SetCellValue("Cargo");

                    ICell cellIn7 = headerRow.CreateCell(6);
                    cellIn7.SetCellValue("Eliminado");

                    var clientes = new List<Cliente>();

                    foreach (var orden in listaOrdenes.data)
                    {
                        var client = new Cliente();
                        client.IdEmpleado = Int32.Parse(orden.IdEmpleado);
                        client.Dni = orden.DNI;
                        client.Nombres = orden.Nombres;
                        client.Planilla = orden.Planilla;
                        client.Area = orden.Area;
                        client.Cargo = orden.Cargo;
                        client.Eliminado = 0;

                        clientes.Add(client);
                    }

                    var clientesDistinc = clientes.DistinctBy(c => c.Dni).ToList();
               
                    for (int i = 0; i < clientesDistinc.Count; i++)
                    {
                        var cliente = new Cliente();
                        IRow rowIn = excelSheet.CreateRow(i+1);

                        ICell _cellIn1 = rowIn.CreateCell(0);
                        _cellIn1.SetCellValue(clientesDistinc[i].IdEmpleado);

                        ICell _cellIn2 = rowIn.CreateCell(1);
                        _cellIn2.SetCellValue(clientesDistinc[i].Dni);

                        ICell _cellIn3 = rowIn.CreateCell(2);
                        _cellIn3.SetCellValue(clientesDistinc[i].Nombres);

                        ICell _cellIn4 = rowIn.CreateCell(3);
                        _cellIn4.SetCellValue(clientesDistinc[i].Planilla);

                        ICell _cellIn5 = rowIn.CreateCell(4);
                        _cellIn5.SetCellValue(clientesDistinc[i].Area);

                        ICell _cellIn6 = rowIn.CreateCell(5);
                        _cellIn6.SetCellValue(clientesDistinc[i].Cargo);

                        ICell _cellIn7 = rowIn.CreateCell(6);
                        _cellIn7.SetCellValue(clientesDistinc[i].Eliminado);
                    }

                workbook.Write(fs);
                fs.Close();
            }
            catch (Exception ex)
            {

                throw;
            }
            


        }

        public void InsertCient(ClienteInput input,IRow rowIn, int id)
        {
            ICell cellIn1 = rowIn.CreateCell(0);
            cellIn1.SetCellValue(id+1);

            ICell cellIn2 = rowIn.CreateCell(1);
            cellIn2.SetCellValue(input.Dni);

            ICell cellIn3 = rowIn.CreateCell(2);
            cellIn3.SetCellValue(input.Nombres);

            ICell cellIn4 = rowIn.CreateCell(3);
            cellIn4.SetCellValue(input.Planilla);

            ICell cellIn5 = rowIn.CreateCell(4);
            cellIn5.SetCellValue(input.Area);

            ICell cellIn6 = rowIn.CreateCell(5);
            cellIn6.SetCellValue(input.Cargo);

            ICell cellIn7 = rowIn.CreateCell(6);
            cellIn7.SetCellValue(0);

        }
    }
}
