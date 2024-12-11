using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CrystalAutomation
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
            /*
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);
            */

            if (Environment.UserInteractive) // Ejecutar en modo consola
            {
                Console.WriteLine("Ejecutando en modo consola para pruebas...");
                Service1 service = new Service1();
                service.TestRun(); // Ejecutar pruebas
                Console.WriteLine("Presiona Enter para salir...");
                Console.ReadLine();
            }
            else // Ejecutar como servicio de Windows
            {
                ServiceBase[] servicesToRun = new ServiceBase[]
                {
                    new Service1()
                };
                ServiceBase.Run(servicesToRun);
            }
        }
    }
}
