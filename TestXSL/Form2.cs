using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestXLS
{
    public partial class Form2 : Form
    {
        public SyIntegradores? Integradores;
        public string? FilePath { get; private set; }
        public string tipoArchivo;
        public Form2()
        {
            InitializeComponent();
            nombre_Box.SelectedIndex = nombre_Box.FindStringExact("durkopp");
            tipoArchivo = nombre_Box.SelectedItem.ToString();
        }

        private void b_cargar_Click(object sender, EventArgs e)
        {
            //Si durkopp está seleccionado, nombrePLC debe tener datos, y si tiene datos, se debe seleccionar un archivo un archivo cuyo nombre contenga nombrePLC
            if (tipoArchivo == "durkopp" && nombrePLC.Text == "")
            {
                MessageBox.Show("Debe ingresar un nombre de PLC");
                return;
            }
                
            string searchString = nombrePLC.Text.Trim(); // Usar el texto ingresado en el TextBox
            string directoryPath = @"C:\mv";

            // Revisa si el directorio existe
            if (Directory.Exists(directoryPath))
            {
                // Busca archivos que contengan el texto ingresado en el nombre
                string[] files = Directory.GetFiles(directoryPath, $"*{searchString}*", SearchOption.TopDirectoryOnly);

                if (files.Any())
                {
                    // Archivo encontrado
                    FilePath = files.First();
                    MessageBox.Show($"Archivo encontrado: {FilePath}", "Archivo Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    // Archivo no encontrado
                    MessageBox.Show("No hay archivo con ese nombre.", "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                // Directorio no encontrado
                MessageBox.Show("La carpeta no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            
        }

        private void nombre_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            tipoArchivo = nombre_Box.SelectedItem.ToString();
        }
    }
}