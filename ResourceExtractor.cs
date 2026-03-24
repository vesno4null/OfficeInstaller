using System;
using System.IO;
using System.Reflection;

namespace OfficeInstaller
{
    public static class ResourceExtractor
    {
        public static string ExtractSetup()
        {
            string temp = Path.Combine(Path.GetTempPath(), "OfficeInstaller");

            if (!Directory.Exists(temp))
                Directory.CreateDirectory(temp);

            string setupPath = Path.Combine(temp, "setup.exe");

            if (!File.Exists(setupPath))
            {
                var assembly = Assembly.GetExecutingAssembly();

                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.setup.exe"))
                using (FileStream file = new FileStream(setupPath, FileMode.Create))
                {
                    stream.CopyTo(file);
                }
            }

            return setupPath;
        }
    }
}