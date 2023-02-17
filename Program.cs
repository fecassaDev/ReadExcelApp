using OfficeOpenXml;
using System.IO;
public class Program
{

    public static void Main()
    {
        ShoWithNoMappings();
        ShoWithMappings();
    }

    private static void ShoWithNoMappings()
    {
        using (var stream = File.Open(@"C:\Users\ferna\Downloads\Carros.xlsx", FileMode.Open))
        {
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var selection = worksheet.Cells["A1:B5"].ToCollection<Carro>(options => options.HeaderRow = 0);

                Console.WriteLine("Exemplo sem mapeamento de propriedades");
                Console.WriteLine(selection.Where(x => x.Modelo == "Fiesta").Select(x => x).First().Modelo);
            }
        }
    }

    private static void ShoWithMappings()
    {
        using (var stream = File.Open(@"C:\Users\ferna\Downloads\Carros.xlsx", FileMode.Open))
        {
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var selection = worksheet.Cells["A1:B5"].ToCollectionWithMappings<Carro>(row =>
                {
                    var carro = new Carro();
                    carro.Marca = row.GetValue<string>("Marca");
                    carro.Modelo = row.GetValue<string>("Modelo");

                    return carro;
                }, options => options.HeaderRow = 0);

                Console.WriteLine("Exemplo com mapeamento de propriedades");
                Console.WriteLine(selection.Where(x => x.Modelo == "Gol").Select(x => x).First().Modelo);

            }
        }
    }
}