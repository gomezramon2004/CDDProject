using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CDDProject
{
    public class MyCsvClassMap : ClassMap<MyDataModel>
    {
        public MyCsvClassMap()
        {
            Map(m => m.Nombre).Name("Nombre del difunto");
            Map(m => m.Fecha).Name("Fecha").TypeConverterOption.Format("dd-MM-yyyy");
            Map(m => m.Bloque).Name("Bloque");
            Map(m => m.Manzana).Name("Manzana");
            Map(m => m.Lote).Name("Lote");
        }
    }

    public class MyDataModel
    {
        public string Nombre { get; set; }
        public string Fecha { get; set; }
        public string Bloque { get; set; }
        public int Manzana { get; set; }
        public string Lote { get; set; }
    }

    public partial class mainprog : Form
    {
        string[] files;
        static string inputFolder;
        public mainprog()
        {
            InitializeComponent();
        }

        static void ConvertToExcel(string[] inputFiles, string outputFolder)
        {
            
            List<MyDataModel> records = new List<MyDataModel>();

            foreach (string inputFile in inputFiles)
            {
                using (StreamReader reader = new StreamReader(inputFile))
                using (CsvReader csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = false }))
                {
                    records.AddRange(csv.GetRecords<MyDataModel>().ToList());
                }
            }

            string outputFilePath = Path.Combine(outputFolder, $"{inputFolder}.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var ExcelPkg = new ExcelPackage())
            {
                ExcelWorksheet ExcelSheet = ExcelPkg.Workbook.Worksheets.Add("Registro de Tumbas");
                ExcelSheet.Cells.LoadFromCollection(records, true);
                ExcelPkg.SaveAs(new FileInfo(outputFilePath));
            }
        }

        private void excelEvent(object sender, EventArgs e)
        { 
            FolderBrowserDialog folderSaveDialog = new FolderBrowserDialog();
            folderSaveDialog.Description = "Selecciona dónde deseas guardar los archivos.";

            if (folderSaveDialog.ShowDialog() == DialogResult.OK)
            {
                ConvertToExcel(files, folderSaveDialog.SelectedPath);
                System.Windows.Forms.MessageBox.Show("¡Se han creado exitosamente los archivos!");
            }

        }

        private void fileEvent(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Selecciona la carpeta que deseas convertir.";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string dirPath = folderBrowserDialog.SelectedPath;
                files = Directory.GetFiles(dirPath, "*.txt", SearchOption.AllDirectories);
                if (files.Length == 0) {
                    System.Windows.Forms.MessageBox.Show("No se encontraron archivos de texto dentro de la carpeta.");
                } else
                {
                    inputFolder = Path.GetFileName(dirPath);
                    fileTxtb.Text = dirPath;
                    excelBtn.Enabled = true;
                }
            }
        }

    }
}
