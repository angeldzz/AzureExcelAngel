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
    private byte[] GenerarGraficoPng
        (Dictionary<string, int> totalesPorCliente)
    {
        var plt = new ScottPlot.Plot();

        double[] valores = totalesPorCliente.Values.Select(v => (double)v).ToArray();
        string[] etiquetas = totalesPorCliente.Keys.ToArray();

        var pie = plt.Add.Pie(valores);
        pie.SliceLabelDistance = 1.3;

        // Asignar etiquetas a cada slice
        for (int i = 0; i < pie.Slices.Count; i++)
        {
            pie.Slices[i].Label = $"{etiquetas[i]}\n({valores[i]:N0}€)";
            pie.Slices[i].LegendText = etiquetas[i];
        }

        plt.Title("Total de Precio por Cliente");
        plt.ShowLegend();
        plt.Axes.Frameless();
        plt.HideAxesAndGrid();

        return plt.GetImageBytes(600, 400, ScottPlot.ImageFormat.Png);
    }
    [Function(nameof(Function1))]
    [BlobOutput("graficos-excel/{name}.png", Connection = "StorageAngel")]
    public async Task Run([BlobTrigger("excel-files/{name}", Connection = "StorageAngel")] BinaryData blobData, string name, Stream outputBlob)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Angel");
        //CREAMOS PACKAGE PARA LEER EL XLSX 
        using var stream = blobData.ToStream();
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
                //CREAMOS UN NUEVO DICCIONARIO
        var totalesPorCliente = new Dictionary<string, int>();
 
        // Generar PNG y subir a Blob
        byte[] png = GenerarGraficoPng(totalesPorCliente);
        await outputBlob.WriteAsync(png);
 
        _logger.LogInformation($"Gráfico PNG guardado en: graficos-output/{name}.png");
  
        cn.Close();
    }

}