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
            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString();
            string previousMonth = ObtenerMesPrevio();
            string nameFile = "LIQUIDACION"+previousMonth + "-" + month+".xlsx";
            string path = Constantes.rutaLiquidaciones +year+"/"+nameFile;
            
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

        //obtener archivo a Escribir

        public ISheet ObtenerExcelParaRegistrar()
        {
            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString();
            string previousMonth = ObtenerMesPrevio();
            string nameFile = "LIQUIDACIONDEMO1.xlsx";
            string path = Constantes.rutaLiquidaciones + year + "/" + nameFile;

            ISheet worksheet = null;

            try
            {
                if (System.IO.File.Exists(path))
                {
                    IWorkbook workbook = null;


                    //FileMode.Create, FileAccess.Write This combination doesnt support for write new data and remove all data
                    using (FileStream FS = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS); 
                        worksheet = workbook.GetSheetAt(0);
                    }

                    //XSSFWorkbook wb1 = null;
                    //using (var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                    //{
                    //    wb1 = new XSSFWorkbook(file);                        
                    //}
                    //wb1.GetSheetAt(0).GetRow(wb1.Count+1).GetCell(0).SetCellValue("Sample");

                    //using (var file2 = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
                    //{
                    //    wb1.Write(file2);
                    //    file2.Close();
                    //}


                    return worksheet;
                }
                


                return worksheet;
            }
            catch (Exception)
            {
                throw;   
            }                         
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
