using Autodesk.DataExchange.Models;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace AU2024_smart_parameter_updater
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Task.Run(async () =>
            {
                try
                {
                    //AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                    Console.Clear();
                    Console.WriteLine("STEP 1 - Read demo configuration from init.json");

                    string fileName = "init.json";
                    string jsonString = File.ReadAllText(fileName);
                    DemoConfiguration cfg = JsonSerializer.Deserialize<DemoConfiguration>(jsonString);

                    Console.WriteLine("STEP 2 - Reading extended data from Excel");
                    DataTable table = ExcelHelper.ReadExcelToDataTable(cfg.Excel);

                    //currently require ADMIN permissions in Visual Studio
                    Console.WriteLine("STEP 3 - Connect to the Autodesk cloud");
                    DataExchangeHelper dxh = new DataExchangeHelper();
                    dxh.Connect(cfg);
                    
                    Console.WriteLine("STEP 4 - Reading data from an existing Data Exchange");
                    ExchangeData exData = await dxh.ReadDataExchange(cfg.ExchangeFileUrn, cfg.ClassName);
                    
                    Console.WriteLine("STEP 5 - Create a new Exchange in ACC");
                    dxh.SetFolder(cfg.ProjectId, cfg.FolderUrn);
                    ExchangeDetails exDetails = await dxh.CreateExchange(cfg.NewExchangeName, "DX SDK demo exchange for AU2024");
                    
                    Console.WriteLine("STEP 6 - Populate the new exchange with data, augmenting the dataset with the excel extended data");
                    await dxh.AddElementsToExchange(exData, exDetails, cfg.ClassName, table).ConfigureAwait(false);

                    Console.WriteLine("\nPress enter to exit...");
                    Console.ReadLine();

                }
                catch (Exception a)
                {
                    Console.WriteLine("An error occurred: " + a);
                    Console.ReadKey();
                }
                finally
                {

                }
            }).GetAwaiter().GetResult();
        }

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var dllName = GetAssemblyName(args) + ".dll";
                var currentAssemblyPath = new System.Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath;
                currentAssemblyPath = Path.GetDirectoryName(currentAssemblyPath);
                if (currentAssemblyPath != null && File.Exists(Path.Combine(currentAssemblyPath, dllName)))
                {
                    return Assembly.LoadFile(Path.Combine(currentAssemblyPath, dllName));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return null;
        }

        private static string GetAssemblyName(ResolveEventArgs args)
        {
            var name = args.Name.IndexOf(",", StringComparison.Ordinal) > -1 ? args.Name.Substring(0, args.Name.IndexOf(",", StringComparison.Ordinal)) : args.Name;
            return name;
        }
    }
}
