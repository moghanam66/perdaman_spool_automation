using System;
using System.IO;
using Autodesk.Navisworks.Api.Automation;

namespace testing_automation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string sourceDirectory = @"C:\Users\11\Desktop\auto_test";
            string resultsDirectory = @"C:\Users\11\Desktop\auto_test\ll";

            // Ensure the results directory exists
            if (!Directory.Exists(resultsDirectory))
            {
                Directory.CreateDirectory(resultsDirectory);
            }

            // Get all .nwd files in the source directory
            string[] nwdFiles = Directory.GetFiles(sourceDirectory, "*.nwd");

            foreach (var file in nwdFiles)
            {
                NavisworksApplication navisworksApplication = null;
                try
                {
                    Console.WriteLine($"Processing file: {file}");
                    navisworksApplication = new NavisworksApplication();
                    navisworksApplication.OpenFile(file);

                    // Call the Examiner Plugin
                    Console.WriteLine("Executing plugin...");
                    navisworksApplication.ExecuteAddInPlugin("ItemInfoPlugin.SAIPEM", "ID_CMD_SP3DSpoolIdent");

                    // Define output file path
                    string outputFileName = Path.Combine(resultsDirectory, Path.GetFileName(file));
                    Console.WriteLine($"Saving file: {outputFileName}");
                    navisworksApplication.SaveFile(outputFileName);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error processing {file}: {e.Message}");
                }
                finally
                {
                    if (navisworksApplication != null)
                    {
                        navisworksApplication.Dispose();
                        Console.WriteLine("Navisworks application disposed.");
                    }
                }
            }
        }
    }
}
