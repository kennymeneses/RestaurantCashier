using DataAccess;
using DataAccess.Modelo;
using DataAccess.Respuestas;
using DataAccess.Solicitudes;
using NPOI.SS.UserModel;

namespace LogicaPlataforma
{
    public class ManejadorClientes : IManejadorClientes
    {
        public ManejadorClientes()
        {

        }

        public Task<RespuestaClienteCreado> CrearCliente(ClienteInput input)
        {
            throw new NotImplementedException();
        }

        public Task<RespuestaClienteEspecifico> ObtenerCliente(string dni)
        {
            throw new NotImplementedException();
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
                cliente.IdEmpleado = row.Cells[0].StringCellValue.Trim();
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
    }
}
