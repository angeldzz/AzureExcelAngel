using Microsoft.Azure.Functions.Worker;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Reflection.Metadata;
using System.Threading.Tasks;

namespace AzureExcelAngel;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function(nameof(Function1))]
    public async Task Run([BlobTrigger("excel-files/{name}", Connection = "StorageAngel")] Stream stream, string name)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Angel");
        //CREAMOS PACKAGE PARA LEER EL XLSX 
        ExcelPackage package = new ExcelPackage(stream);
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        string connectionString = @"Data Source=LOCALHOST\DEVELOPER;Initial Catalog=HOSPITAL;Persist Security Info=True;User ID=sa;Encrypt=True;Trust Server Certificate=True";
        string sql = "insert into FACTURASAZURE (CLIENTE, CONCEPTO, PRECIO) values "
        + " (@cliente, @concepto, @precio)";
        SqlConnection cn = new SqlConnection(connectionString);
        SqlCommand com = new SqlCommand();
        com.Connection = cn;
        com.CommandType = CommandType.Text;
        com.CommandText = sql;
        cn.Open();
        //EXCEL COMIENZA SUS FILAS EN 1 (NUESTRA CABECERA) 
        for (int i = 2; i <= worksheet.Dimension.Rows; i++)
        {

            string cliente = worksheet.Cells[i, 1].Text;

            string concepto = worksheet.Cells[i, 2].Text;

            int precio = int.Parse(worksheet.Cells[i, 3].Text);

            com.Parameters.AddWithValue("@cliente", cliente);

            com.Parameters.AddWithValue("@concepto", concepto);

            com.Parameters.AddWithValue("@precio", precio);

            com.ExecuteNonQuery();

            com.Parameters.Clear();

            _logger.LogInformation("Cliente: " + cliente);

        }



        cn.Close();
    }
}