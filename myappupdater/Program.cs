// A minimal console app or WinForms app called "MyApp.Updater.exe"
using System;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace MyAppUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: MyApp.Updater.exe <NewExePath> <OldExePath>");
                return;
            }

            string newExe = args[0];
            string oldExe = args[1];

            // 1. Wait for old process to exit
            //    (Optionally, you could pass the process ID or name to kill or wait)
            Thread.Sleep(3000); // A simple wait.

            try
            {
                // 2. Replace old file with new
                File.Copy(newExe, oldExe, true);

                // 3. Optionally delete the downloaded file from temp
                File.Delete(newExe);

                // 4. Restart the new app
                Process.Start(oldExe);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error updating file: " + ex.Message);
            }
        }
    }
}
