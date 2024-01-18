using System;
using Python.Runtime;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CDDProject

{

    public partial class mainprog : Form
    {
        string[] files;
        public mainprog()
        {
            InitializeComponent();
            Runtime.PythonDLL = @"C:\Users\kapig\AppData\Local\Programs\Python\Python312\python312.dll";
        }

        private void excelEvent(object sender, EventArgs e)
        {
            FolderBrowserDialog folderSaveDialog = new FolderBrowserDialog();
            folderSaveDialog.Description = "Selecciona dónde deseas guardar los archivos.";

            if (folderSaveDialog.ShowDialog() == DialogResult.OK)
            {
                PythonEngine.Initialize();
                dynamic pythonScript = Py.Import("testCDD");
                using (Py.GIL())
                {
                    dynamic outputFolder = new PyString(folderSaveDialog.SelectedPath);

                    foreach (string file in files)
                    {
                        dynamic inputFile = new PyString(file);
                        dynamic result = pythonScript.InvokeMethod("convertToExcel", new PyObject[] { inputFile, outputFolder });
                    }

                    System.Windows.Forms.MessageBox.Show("¡Se han creado exitosamente los archivos!");
                }
            }

        }

        private void fileEvent(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Selecciona la carpeta que deseas convertir.";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string dirPath = folderBrowserDialog.SelectedPath;
                if (Directory.GetFiles(dirPath, "*.txt").Length == 0) {
                    System.Windows.Forms.MessageBox.Show("No se encontraron archivos de texto dentro de la carpeta.");
                } else
                {
                    files = Directory.GetFiles(folderBrowserDialog.SelectedPath, "*.txt", SearchOption.AllDirectories);
                    fileTxtb.Text = dirPath;
                    excelBtn.Enabled = true;
                }
            }
        }

    }
}
