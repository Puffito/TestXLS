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
        public string PLC { get; private set; }
        public Form2()
        {
            InitializeComponent();
            nombre_Box.SelectedIndex = nombre_Box.FindStringExact("durkopp");
            tipoArchivo = nombre_Box.SelectedItem.ToString();
        }

        private void b_cargar_Click(object sender, EventArgs e)
        {
            //nombrePLC no puede ser vacio
            if (nombrePLC.Text == "")
            {
                MessageBox.Show("Debe ingresar un nombre de PLC");
                return;
            }
                
            string searchString = nombrePLC.Text.Trim(); // Usar el texto ingresado en el TextBox
            
            //Seleccionar archivo Excel a cargar
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Seleccionar archivo Excel";
            openFileDialog.ShowDialog();
            FilePath = openFileDialog.FileName;
            if (FilePath == "")
            {
                MessageBox.Show("Debe seleccionar un archivo Excel");
                return;
            }

            //Devolver OK y cerrar ventana
            PLC = nombrePLC.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
                       
        }

        private void nombre_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            tipoArchivo = nombre_Box.SelectedItem.ToString();
            if(nombre_Box.SelectedItem.ToString() == "tgw")
            {
                textoTipo.Text = "Prefijo clase";
            }
            else
            {
                textoTipo.Text = "Nombre PLC";
            }
        }
    }
}