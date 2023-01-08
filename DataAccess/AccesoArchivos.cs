using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataAccess
{
    public class AccesoArchivos
    {
        public AccesoArchivos()
        {

        }

        public ISheet ObtenertHistorialdeLiquidacionActual()
        {
            string nameFile = "LIQUIDACION" + GetMonths() + ".xlsx";
            string path = Constantes.rutaLiquidaciones +GetYear()+"/"+nameFile;
            
            ISheet worksheet = null;
            try
            {
                if(System.IO.File.Exists(path))
                {
                    IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              

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
            //validacion si hoy es mayor al dia 15
            // esto debe entregar por ejemplo 12-1 o 1-2

            //validar tb el año

            int day = DateTime.Now.Day;
            int december = 12;
            int month = DateTime.Now.Month;

            string response = string.Empty;

            if (month == 1 && day <= 15)
            {
                response = december.ToString() + "-" + month.ToString();
            }

            else if (day > 15)
            {
                response = month.ToString() + "-" + (month + 1).ToString();
            }

            else if (day <= 15)
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


            if (day > 15 && month == 12)
            {
                year = year + 1;
            }
            return year.ToString();
        }

        private string ObtenerMesPrevio()
        {
            int lastMonth = 12;
            int month = DateTime.Now.Month;
            
            if(month == 1)
            {
                return lastMonth.ToString();
            }

            lastMonth = month -1;

            return lastMonth.ToString();
        }
    }
}
