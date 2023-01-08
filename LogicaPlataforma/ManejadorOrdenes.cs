using DataAccess;
using DataAccess.Modelo;
using DataAccess.Respuestas;
using DataAccess.Solicitudes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace LogicaPlataforma
{
    public class ManejadorOrdenes : IManejadorOrdenes
    {
        AccesoArchivos _accesoArchivos;

        public ManejadorOrdenes(AccesoArchivos accesoArchivos)
        {
            _accesoArchivos = accesoArchivos;
        }

        public Task ObtenerMesesFacturados()
        {
            throw new NotImplementedException();
        }

        public async Task<RespuestaFacturacionMesActual> ObtenerFacturacionMesActual()
        {
            var sheetMesActual = _accesoArchivos.ObtenertHistorialdeLiquidacionActual();
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
                response.Total = sheetMesActual.LastRowNum;
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

            //var sheetMesActual = _accesoArchivos.ObtenerExcelParaRegistrar();

            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString();
            string nameFile = "LIQUIDACION" + ObtenerMesPrevio() + "-" + month + ".xlsx";
            string path = Constantes.rutaLiquidaciones + year + "/" + nameFile;

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
                    
                        IRow rowIn = excelSheet.CreateRow(numRecords+1);
                        InsertValuesInNewRow(order, rowIn, LastId);
                    
                        var file = new FileStream(path, FileMode.Create);
                        workbook.Write(file);

                        file.Close();

                        respuesta.Id = LastId + 1;
                        respuesta.ResponseStatus = Constantes.StatusOk;
                        respuesta.ResponseDescription = Constantes.OrderRegisteredSuccessfully;
                    }
                }
                else
                {
                    
                    string _path = Constantes.rutaLiquidaciones+year+"/";

                    FileStream fs = new FileStream(path, FileMode.Create);

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Hoja1");
                    IRow headerRow = excelSheet.CreateRow(0);
                    IRow rowIn = excelSheet.CreateRow(1);
                    InsertHeaders(headerRow);
                    InsertValuesInNewRow(order, rowIn, LastId);

                    //ICell cellIn1 = rowIn.CreateCell(0);
                    //cellIn1.SetCellValue(year);

                    workbook.Write(fs);

                    fs.Close();

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

            ICell cellIn18 = rowIn.CreateCell(16);
            cellIn18.SetCellValue("No");
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
                if(!Directory.Exists(path+"nuevacarpeta"))
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

                    respuesta.ResponseDescription = Constantes.OrderRegisteredSuccessfully;
                    respuesta.ResponseStatus = Constantes.StatusOk;
                    respuesta.Id = 1;

                    return respuesta;
                }

                return respuesta;
            }
            catch (Exception ex)
            {
                respuesta.ResponseStatus= Constantes.StatusError;
                respuesta.ResponseDescription = ex.Message;
                return respuesta;
            }            
        }

        private string ObtenerMesPrevio()
        {
            //validacion si hoy es mayor al dia 15

            int lastMonth = 12;
            int month = DateTime.Now.Month;

            if (month == 1)
            {
                return lastMonth.ToString();
            }

            lastMonth = month - 1;

            return lastMonth.ToString();
        }
    }
}
