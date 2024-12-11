using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

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

            string inputFolder = @"C:\Users\Soporte2\Desktop\FOLDER_INPUT"; // Cambia esta ruta
            string outputFolder = @"C:\Users\Soporte2\Desktop\FOLDER_OUTPUT"; // Cambia esta ruta

            // Crear la carpeta de salida si no existe
            if (!Directory.Exists(outputFolder))
            {
                Console.WriteLine("Creando carpeta OUTPUT");
                Directory.CreateDirectory(outputFolder);
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

            Console.WriteLine("Crystal Report iniciado ...");
            EventLog.WriteEntry("CrystalReportService iniciado.");

            string pdfOutputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(inputPath) + ".pdf");
            
            try
            {

                Console.WriteLine("Cargando archivo .rpt...");
                ReportDocument reportDocument = new ReportDocument();
                reportDocument.Load(inputPath);

                Console.WriteLine("Exportando el archivo a PDF...");
                reportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, pdfOutputPath);

                Console.WriteLine("Modificando el archivo PDF...");
                ModifyPdf(pdfOutputPath, pdfOutputPath, "89231700-3");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al procesar archivo: {ex.Message}");
            }
        }

        // Crear una fuente estándar (Helvetica, Times-Roman, etc.)
        //PdfFont font = PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA);

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
