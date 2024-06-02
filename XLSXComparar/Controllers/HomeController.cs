using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Globalization;

namespace XLSXComparar.Controllers
{
    [ApiController]
    [Route("api/[Controller]")]
    public class HomeController : Controller
    {
        public HomeController()
        {
            // Establecer el contexto de la licencia
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        [HttpPost("Actualizar")]
        public IActionResult ActualizarFile(List<IFormFile> files) 
        {
            if (files == null || !files.Any())
            {
                return BadRequest("No se subio el archivo");
            }

            var PrimerArchivo = files.First();
            var ArchivosAComparar = files.Skip(1).ToList();

            try { 

            using (var Flujo = new MemoryStream())
            {
                PrimerArchivo.CopyTo(Flujo);
                using (var Paquete = new ExcelPackage(Flujo))
                {
                    var Worksheet = Paquete.Workbook.Worksheets.First();
                    var PrimerData = GetDataFromWorksheet(Worksheet);

                    foreach (var ArchivoAComparar in ArchivosAComparar) 
                    {
                        try
                        {
                            using (var ComparisonStream = new MemoryStream())
                            {
                                ArchivoAComparar.CopyTo(ComparisonStream);
                                using (var comparisonPackage = new ExcelPackage(ComparisonStream))
                                {
                                    var comparisonWorksheet = comparisonPackage.Workbook.Worksheets.First();
                                    var comparisonData = GetDataFromWorksheet(comparisonWorksheet);
                                    CompararAndActualizar(PrimerData, comparisonData);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            return BadRequest($"Error procesado en el archivo {ArchivoAComparar.FileName}: {ex.Message} ");
                        }
                    }
                    ActualizarWorksheet(Worksheet, PrimerData);
                    return File(Paquete.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "GRISELUP.xlsx");
                    
                }
            }
        }
        catch (Exception ex)
        {
                return BadRequest($"Error procesando el archivo principal {PrimerArchivo.FileName}: {ex.Message}");
        }
    }








private List<Dictionary<string, string>> GetDataFromWorksheet(ExcelWorksheet worksheet)
        {
            var data = new List<Dictionary<string, string>>();
            var rowCount = worksheet.Dimension.Rows;
            var ColCount = worksheet.Dimension.Columns;
            var headers = new List<string>();

            for (int col = 1; col <= ColCount; col++)
            {
                var header = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                headers.Add(header);
            }

            for (int row = 2; row <= rowCount; row++) 
            {
                var rowData = new Dictionary<string, string>();
                for (int col = 1; col <= ColCount; col++)
                {
                    var columnName = headers[col - 1];
                    rowData[columnName] = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                }
                data.Add(rowData);
            }
            return data;
        }






        private void CompararAndActualizar(List<Dictionary<string, string>> PrimerArchivo, List<Dictionary<string, string>> ComparacionData)
        {
            foreach (var PrimerRow in PrimerArchivo)
            {

                var match = ComparacionData.FirstOrDefault(ComparacionData =>
                ComparacionData["nombre"] == Capitalize(PrimerRow["nombre"])&&
                ComparacionData["apellido paterno"] == Capitalize(PrimerRow["apellido paterno"]) &&
                ComparacionData["apellido materno"] == Capitalize(PrimerRow["apellido materno"]));


                if (match != null)
                {
                    PrimerRow["casilla"] = match.ContainsKey("casilla") ? match["casilla"] : string.Empty;
                    PrimerRow["consecutivo"] = match.ContainsKey("consecutivo")  ? match["consecutivo"] :  string.Empty;
                    PrimerRow["seccional"] = match.ContainsKey("seccional") ? match["seccional"] :  string.Empty;
                }

            }
        }






        private void ActualizarWorksheet(ExcelWorksheet worksheet, List<Dictionary<string, string>> data)
        {
            var rowCount = data.Count;
            var colCount = worksheet.Dimension.Columns;
            var headers = new List<string>();

            for (int col = 1; col <= colCount; col++)
            {
                var header = worksheet.Cells[1, col].Value?.ToString()?.Trim() ?? $"Column{col}";
                headers.Add(header);
            }

            for (int row = 2; row <= rowCount + 1; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    var columnName = headers[col - 1];
                    if (data[row - 2].ContainsKey(columnName))
                    {
                        worksheet.Cells[row, col].Value = data[row - 2][columnName];
                    }
                }
            }
        }


        private string Capitalize(string input)
        {
            if(string.IsNullOrWhiteSpace(input))
                return string.Empty;

            input = input.ToLower();
            TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(input);
        }

    }
}
