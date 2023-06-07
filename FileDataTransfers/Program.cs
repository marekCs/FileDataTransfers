using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeOpenXml;

internal class Program
{
    private static void Main(string[] args)
    {

        // Nastavení licence pro EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Cesta k souboru Excel
        var excelFilePath = @"C:\Users\marek\Documents\airradio\RADIOFONIA.xlsx";

        // Cesta k XML souborům
        var xmlFolderPath = @"C:\Users\marek\Documents\airradio\BASI";

        // Cesta k složce, kam se mají přesunout zpracované soubory
        var parsedFolderPath = @"C:\Users\marek\Documents\airradio\parsed\BASI";
        if (!Directory.Exists(parsedFolderPath))
            Directory.CreateDirectory(parsedFolderPath); // vytvoříme složku, pokud neexistuje

        using (var excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = excelPackage.Workbook.Worksheets[0]; // předpokládáme, že data jsou na prvním listu

            // Projděte všechny soubory ve složce
            foreach (var xmlFilePath in Directory.GetFiles(xmlFolderPath, "*.xml"))
            {
                // Načtěte XML soubor
                var xmlDocument = XDocument.Load(xmlFilePath);

                // Získejte potřebná data z XML souboru
                var dbTupleElement = xmlDocument.Root.Element("DbTuple");
                var title = dbTupleElement.Elements().FirstOrDefault(e => e.Name.LocalName == "STRING_1")?.Value;
                var artist = dbTupleElement.Elements().FirstOrDefault(e => e.Name.LocalName == "STRING_2")?.Value;
                var author = dbTupleElement.Elements().FirstOrDefault(e => e.Name.LocalName == "USER_REC")?.Value;
                var filename = dbTupleElement.Elements().FirstOrDefault(e => e.Name.LocalName == "FILE")?.Value;

                // Najděte odpovídající řádek v Excelu
                var rowNumber = 1;
                var totalRows = worksheet.Dimension?.Rows ?? 1; // získejte počet řádků nebo nastavte na 1, pokud není definováno
                while (rowNumber <= totalRows && (worksheet.Cells[rowNumber, 1].Value == null || !worksheet.Cells[rowNumber, 1].Value.ToString().EndsWith(".mp3") || worksheet.Cells[rowNumber, 1].Value.ToString() != filename))
                    rowNumber++;

                // Pokud byl nalezen odpovídající řádek, aktualizujte data v Excelu
                if (rowNumber <= totalRows && worksheet.Cells[rowNumber, 1].Value != null)
                {
                    worksheet.Cells[rowNumber, 5].Value = title;  // Title
                    worksheet.Cells[rowNumber, 6].Value = artist; // Artist
                    worksheet.Cells[rowNumber, 7].Value = author; // Author

                    // Přesunout zpracovaný soubor do složky "parsed"
                    var newFilePath = Path.Combine(parsedFolderPath, Path.GetFileName(xmlFilePath));
                    File.Move(xmlFilePath, newFilePath);
                }
            }

            // Uložte změny v Excel souboru
            excelPackage.Save();
        }

        Console.WriteLine("Data byla úspěšně přenesena do Excelu.");
    }
}