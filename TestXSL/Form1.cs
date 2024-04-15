using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using DTutilities;
using System.Xml;
using System.Text;
using System.Text.RegularExpressions;

namespace TestXLS
{
    public partial class Form1 : Form
    {
        string pathApp;
        string filePath;
        string tipoArchivo;
        SyIntegradores Integradores;

        public Form1()
        {
            InitializeComponent();

            Inicializar();
        }

        private void Inicializar()
        {
            //--------- Capturo el directorio de la aplicacion
            pathApp = Environment.CurrentDirectory;

            if (File.Exists(pathApp + "\\cfg.xml"))
            {
                //------------- Cargo los datos
                CargarCfg();
            }
            else
            {
                TSendToLog st1 = new TSendToLog("No existe el archivo de configuración cfg.xml", pathApp);
            }
        }

        private void CargarCfg()
        {
            XmlElement xmlIntegrador;
            XmlDocument xmlArchivo;
            XmlNodeList ListaIntegradores;
            SyIntegrador Integrador;
            int i;

            Integradores = new SyIntegradores();

            xmlArchivo = new XmlDocument();

            xmlArchivo.Load(pathApp + "\\cfg.xml");

            if (xmlArchivo.DocumentElement.Name == "cfg")
            {
                ListaIntegradores = xmlArchivo.GetElementsByTagName("Integrador");

                for (i = 0; i < ListaIntegradores.Count; i++)
                {
                    xmlIntegrador = ListaIntegradores[i] as XmlElement;

                    Integrador = new SyIntegrador(xmlIntegrador.GetAttribute("nombre"), xmlIntegrador.GetAttribute("tipo"), xmlIntegrador.GetAttribute("equipos"),
                        xmlIntegrador.GetAttribute("indice"), xmlIntegrador.GetAttribute("indice2"));

                    Integradores.Listado.Add(Integrador);
                }

                /*
                for (i=0; i < Integradores.Listado.Count; i++)
                {
                    SyIntegrador integ = Integradores.Listado[i] as SyIntegrador;
                }
                */
            }

            TSendToLog st1 = new TSendToLog("Cargados los datos del fichero cfg.xml", pathApp);
        }

        #region Lectura de archivo
        public Dictionary<string, Dictionary<string, List<string>>> ReadDataFromSheet(string filePath, string nombreIntegrador)
        {
            // Encontrar el integrador seleccionado
            SyIntegrador integradorSeleccionado = Integradores.Listado.Find(x => x.nombre == nombreIntegrador);

            // Extraer las celdas de inicio de los datos del integrador seleccionado
            string componentNameCell = integradorSeleccionado.tipo;
            string nameValuesStartCell = integradorSeleccionado.equipos;
            string indexStartCell = integradorSeleccionado.indice;
            string secondaryIndexStartCell = integradorSeleccionado.indice2;

            // Convertir las celdas de inicio de los datos a índices de fila y columna
            int nameValuesStartRow = int.Parse(nameValuesStartCell.Substring(1));
            int nameValuesStartColumn = calcularLetra(nameValuesStartCell.Substring(0, 1));
            int indexStartRow = int.Parse(indexStartCell.Substring(1));
            int indexStartColumn = calcularLetra(indexStartCell.Substring(0, 1));
            int secondaryIndexStartRow = int.Parse(secondaryIndexStartCell.Substring(1));
            int secondaryIndexStartColumn = calcularLetra(secondaryIndexStartCell.Substring(0, 1));

            // Inicializar un diccionario para almacenar los datos extraídos
            Dictionary<string, Dictionary<string, List<string>>> extractedData = new Dictionary<string, Dictionary<string, List<string>>>();

            // Crear una instancia de la aplicación Excel y abrir el libro de trabajo
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath, ReadOnly: true);

            // Iterar sobre todas las hojas del libro de trabajo
            foreach (Excel.Worksheet xlWorksheet in xlWorkbook.Worksheets)
            {
                // Mirar si la hoja es visible y si su nombre consiste solo en dígitos
                if (xlWorksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible && xlWorksheet.Name.All(char.IsDigit))
                {
                    // Sacar el nombre del componente de la celda especificada
                    string componentType = xlWorksheet.Range[componentNameCell].Value2?.ToString();

                    // Si el nombre del componente no está en el diccionario, añadirlo
                    extractedData.Add(componentType, new Dictionary<string, List<string>>());

                    // Encontrar la última fila de la hoja de cálculo
                    int lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                   // Inicializar listas para almacenar los nombres, índices y índices secundarios
                    List<string> names = new List<string>();
                    List<string> indexes = new List<string>();
                    List<string> secondaryIndexes = new List<string>();

                    // Extraer los nombres, índices y índices secundarios de las celdas especificadas
                    for (int i = nameValuesStartRow; i <= lastRow; i++)
                    {
                        string nameValue = xlWorksheet.Cells[i, nameValuesStartColumn].Value2?.ToString();
                        if (string.IsNullOrEmpty(nameValue))
                        {
                            break;
                        }
                            names.Add(nameValue);
                            indexes.Add(xlWorksheet.Cells[i, indexStartColumn].Value2?.ToString() ?? "");
                        secondaryIndexes.Add(xlWorksheet.Cells[i, secondaryIndexStartColumn].Value2?.ToString() ?? "");
                    }

                    // Añadir los datos extraídos al diccionario
                    extractedData[componentType].Add("Names", names);
                    extractedData[componentType].Add("Indexes", indexes);
                    extractedData[componentType].Add("SecondaryIndexes", secondaryIndexes);
                }
            }

            // Cerrar el libro de trabajo y la aplicación Excel
            xlWorkbook.Close(false);
            xlApp.Quit();

            // Liberar los recursos COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //Devolver los datos extraídos
            return extractedData;
        }

        #endregion

        #region Guardado de datos
            
        private void WriteToCSV(Dictionary<string, Dictionary<string, List<string>>> extractedData, string excelFileName, bool add, string? fileToAdd = "C:/mv/mcmscadaopc1.csv")
        {
            // Cogemos el nombre del archivo de Excel y sacamos la parte que nos interesa
            string fileName = ExtractFileName(excelFileName);

            // Define el encabezado del archivo CSV
            string header = "CONEXIÓN;AREA;ZONA;PLANTA;EQL;TIPO;TYPE;EQUIPO;POSI;AUX;CLASE;DIRECCIONAMIENTO;RUN;BOX DETECTED;CLASE AUX.";

            // Escritura del encabezado en el archivo CSV
            using (StreamWriter writer = new StreamWriter(fileToAdd, add, Encoding.UTF8))
            {
                if(!add)writer.WriteLine(header);

                // Write data to the CSV file
                foreach (var sheetEntry in extractedData)
                {
                    string componentType = sheetEntry.Key;
                    Dictionary<string, List<string>> sheetData = sheetEntry.Value;

                    List<string> names = sheetData["Names"];
                    List<string> indexes = sheetData["Indexes"];
                    List<string> secondaryIndexes = sheetData["SecondaryIndexes"];


                    for (int i = 0; i < names.Count; i++)
                    {
                        writer.Write(fileName); // CONEXIÓN
                        writer.Write(";");
                        writer.Write(";"); // AREA
                        writer.Write(fileName); // ZONA
                        writer.Write(";");
                        writer.Write(";"); // PLANTA
                        writer.Write(";"); // EQL
                        writer.Write(componentType); // TIPO
                        writer.Write(";");
                        writer.Write(";"); // TYPE
                        writer.Write(fileName + "_" + names[i]); // EQUIPO
                        writer.Write(";");
                        writer.Write(";"); // POSI
                        writer.Write(";"); // AUX
                        writer.Write(";"); // CLASE

                        // Escritura de los valores de los índices y los índices secundarios bajo DIRECCIONAMIENTO
                        string indexValue = indexes[i];
                        string secondaryIndexValue = secondaryIndexes[i];

                        bool indexWritten = false; // Marcador para indicar si el valor del índice ha sido escrito

                        // Revisa que el primer índice no esté vacío y contenga solo dígitos
                        if (!string.IsNullOrEmpty(indexValue) && indexValue.All(char.IsDigit))
                        {
                            writer.Write(indexValue);
                            indexWritten = true; // Marca el índice como escrito
                        }

                        // Revisa que el segundo índice no esté vacío y contenga solo dígitos
                        if (!string.IsNullOrEmpty(secondaryIndexValue) && secondaryIndexValue.All(char.IsDigit))
                        {
                            //Añade un espacio si el índice principal ya ha sido escrito
                            if (indexWritten) writer.Write(" ");

                            writer.Write(secondaryIndexValue);
                        }

                        writer.Write(";"); // RUN
                        writer.Write(";"); // BOX DETECTED
                        writer.WriteLine(";"); // CLASE AUX
                    }
                }
            }
        }
        #endregion

        #region Metodos auxiliares

        private int calcularLetra(string letra)
        {
            int columnNumber = 0;
            int mul = 1;
            for (int i = letra.Length - 1; i >= 0; i--)
            {
                columnNumber += (letra[i] - 'A' + 1) * mul;
                mul *= 26;
            }
            return columnNumber;
        }

        private static string ExtractFileName(string excelFileName)
        {
            // Regex que busca el nombre del archivo de Excel que comienza con "PLC" y termina con "_"
            Match match = Regex.Match(excelFileName, @"PLC[^_]+_");

            if (match.Success)
            {
                string extractedString = match.Value;

                // Quitamos el guion bajo al final
                extractedString = extractedString.Substring(0, extractedString.Length - 1);

                return extractedString;
            }
            else
            {
                // Si no se encuentra ninguna coincidencia, devolvemos una cadena vacía
                return string.Empty;
            }
        }
        #endregion

        #region Botones
        private void b_cargar_Click(object sender, EventArgs e)
        {
            using (var secondForm = new Form2())
            {
                var result = secondForm.ShowDialog();

                secondForm.Integradores = Integradores;

                if (result == DialogResult.OK && !string.IsNullOrEmpty(secondForm.FilePath))
                {
                    filePath = secondForm.FilePath;

                    tipoArchivo = secondForm.tipoArchivo;

                    textBox1.Text = filePath;
                }
            }
        }

        private void b_anadir_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Filter = "Archivo CSV |*.csv";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    string toFile = fileDialog.FileName;
                    WriteToCSV(ReadDataFromSheet(filePath, tipoArchivo), filePath, true, toFile);
                    MessageBox.Show("Archivo guardado correctamente", "Completado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void b_crear_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                string toFile = filePath.Replace(".xlsx", ".csv");
                WriteToCSV(ReadDataFromSheet(filePath, tipoArchivo), filePath, false);
                MessageBox.Show("Archivo guardado correctamente", "Completado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        #endregion
    }
}