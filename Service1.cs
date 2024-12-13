using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

using iTextPdf = iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

namespace CrystalAutomation
{
    public partial class Service1 : ServiceBase
    {
        private FileSystemWatcher watcher;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Console.WriteLine("Iniciando servicio CrystalReportService...");

            Console.WriteLine("Asignando carpetas INPUT y OUTPUT");

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string inputFolder = Path.Combine(desktopPath, "FOLDER_INPUT");
            string outputFolder = Path.Combine(desktopPath, "FOLDER_OUTPUT");
            string erpFolder = Path.Combine(desktopPath, "FOLDER_ERP");

            // Crear la carpeta de salida si no existe
            if (!Directory.Exists(outputFolder))
            {
                Console.WriteLine("Creando carpeta OUTPUT");
                Directory.CreateDirectory(outputFolder);
            }

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine("Creando carpeta OUTPUT");
                Directory.CreateDirectory(inputFolder);
            }

            if (!Directory.Exists(erpFolder))
            {
                Console.WriteLine("Creando carpeta ERP OLD");
                Directory.CreateDirectory(erpFolder);
            }

            Console.WriteLine("Configurando el observador de la carpeta INPUT FOLDER");
            // Configurar el observador de carpeta
            watcher = new FileSystemWatcher
            {
                Path = inputFolder,
                Filter = "*.rpt",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };

            watcher.Created += (sender, e) => ProcessFile(e.FullPath, outputFolder);
            //watcher.Changed += (sender, e) => ProcessFile(e.FullPath, outputFolder);

            watcher.EnableRaisingEvents = true;

        }

        protected override void OnStop()
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
            EventLog.WriteEntry("CrystalReportService detenido.");
        }

        private void ProcessFile(string inputPath, string outputFolder)
        {

            try
            {
                //ReportDocument reportDocument = new ReportDocument();

                using (ReportDocument reportDocument = new ReportDocument())
                {

                    Console.WriteLine("Renombrando archivo en Input...");
                    string renamedInputPath = RenameFileInInput(inputPath);

                    Console.WriteLine("Crystal Report iniciado ...");
                    EventLog.WriteEntry("CrystalReportService iniciado.");

                    Console.WriteLine($"Archivo encontrado: {File.Exists(renamedInputPath)}");
                    Console.WriteLine($"Ruta del archivo: {renamedInputPath}");
                    Console.WriteLine($"Cargando el archivo .rpt...");

                    reportDocument.Load(renamedInputPath);

                    Console.WriteLine("Generando ruta única para Output...");
                    string renamePdfOutputPath = GenerateUniqueOutputPath(renamedInputPath, outputFolder);

                    Console.WriteLine("Exportando el archivo a PDF...");
                    reportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, renamePdfOutputPath);

                    // Validar que el archivo contega el rut 88888888-8
                    Console.WriteLine("Validando contenido del archivo PDF...");
                    bool response = ContainsValueInPdf(renamePdfOutputPath, "88888888-8");

                    if(response == false)
                    {

                       Console.WriteLine("El archivo no contiene el valor esperado. No se realizarán más acciones.");
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Modificando el archivo PDF...");
                        ModifyPdf(renamePdfOutputPath, renamePdfOutputPath, "89231700-3");
                    }

                    // Mover renamedInputPath a otra carpeta erpFolder
                    Console.WriteLine("Moviendo el archivo procesado a la carpeta ERP...");
                    string erpFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "FOLDER_ERP");

                    // Crear la carpeta si no existe
                    if (!Directory.Exists(erpFolder))
                    {
                        Directory.CreateDirectory(erpFolder);
                    }

                    string destinationPath = Path.Combine(erpFolder, Path.GetFileName(renamedInputPath));

                    File.Move(renamedInputPath, destinationPath);
                    Console.WriteLine($"Archivo movido a: {destinationPath}");


                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al procesar archivo: {ex.Message}");
            }
        }

        private void ModifyPdf(string inputPdfPath, string outputPdfPath, string overlayText)
        {
            try
            {
                // Abrir el PDF de entrada
                PdfDocument pdfDocument = PdfReader.Open(inputPdfPath, PdfDocumentOpenMode.Modify);

                foreach (PdfPage page in pdfDocument.Pages)
                {

                    // Crear un objeto gráfico para dibujar sobre la página
                    XGraphics gfx = XGraphics.FromPdfPage(page);

                    // Configurar la fuente y el color del texto
                    XFont font = new XFont("Times New Roman", 8);
                    XBrush brush = XBrushes.Black;

                    // Posición del texto en la página (ajustar según sea necesario)
                    double x = 227; // Coordenada X
                    double y = 43;  // Coordenada Y
                    double width = 40;  // Ancho del cuadro blanco
                    double height = 10;  // Alto del cuadro blanco

                    // Dibujar un rectángulo blanco sobre el área deseada
                    XBrush whiteBrush = XBrushes.White;
                    
                    gfx.DrawRectangle(whiteBrush, x + 310, y, width, height);
                    gfx.DrawRectangle(whiteBrush, x, y, width, height);

                    // Dibujar el texto sobre el cuadro blanco
                    gfx.DrawString(overlayText, font, brush, x + 5, y + 10); // Desplazamos el texto un poco dentro del cuadro
                    gfx.DrawString(overlayText, font, brush, x + 308, y + 10); // Desplazamos el texto un poco dentro del cuadro


                }

                // Guardar el PDF modificado
                pdfDocument.Save(outputPdfPath);
                Console.WriteLine($"Archivo PDF modificado guardado en: {outputPdfPath}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al modificar el PDF: {ex.Message}");
            }
        }

        private bool ContainsValueInPdf(string pdfPath, string value)
        {
            try
            {
                using (iTextPdf.PdfReader reader = new iTextPdf.PdfReader(pdfPath))
                using (iTextPdf.PdfDocument pdfDoc = new iTextPdf.PdfDocument(reader))
                {
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        string pageContent = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i));
                        if (pageContent.Contains(value))
                        {
                            Console.WriteLine($"El valor {value} fue encontrado en la página {i}");
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el PDF: {ex.Message}");
            }

            Console.WriteLine($"El valor {value} no fue encontrado en el PDF.");
            return false;
        }

        private string RenameFileInInput(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            string extension = Path.GetExtension(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            // Generar un nuevo nombre único
            string uniqueName = $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}";

            string newFilePath = Path.Combine(directory, uniqueName);

            // Renombrar el archivo
            File.Move(filePath, newFilePath);
            

            Console.WriteLine($"Archivo renombrado en Input: {newFilePath}");
            return newFilePath;
        }

        private string GenerateUniqueOutputPath(string originalPath, string outputFolder)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
            string uniqueName = $"{fileNameWithoutExtension}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

            string outputPath = Path.Combine(outputFolder, uniqueName);
            Console.WriteLine($"Archivo PDF generado con nombre único: {outputPath}");
            return outputPath;
        }

        public void TestRun()
        {
            OnStart(null); // Simula el inicio del servicio
            Console.WriteLine("El servicio está en ejecución. Presiona 'q' y Enter para detener el servicio...");

            // Mantener el servicio activo hasta que se presione 'q'
            while (Console.ReadLine()?.ToLower() != "q")
            {

            }

            OnStop(); // Simula la detención del servicio
            Console.WriteLine("El servicio simulado se ha detenido.");
        }

    }
    
}
