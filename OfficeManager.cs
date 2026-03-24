using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace OfficeInstaller
{
    public class OfficeManager
    {
        private static void SetOfficeRegistry()
        {
            ProcessStartInfo reg = new ProcessStartInfo
            {
                FileName = "reg",
                Arguments = "add \"HKCU\\Software\\Microsoft\\Office\\16.0\\Common\\ExperimentConfigs\\Ecs\" /v \"CountryCode\" /t REG_SZ /d \"std::wstring|US\" /f",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            Process.Start(reg)?.WaitForExit();
        }

        public static void Install(
            string version,
            bool word,
            bool excel,
            bool ppt,
            bool outlook,
            bool access,
            bool publisher,
            bool lync,
            bool onedrive)
        {
            // Записываем ключ реестра
            SetOfficeRegistry();

            // Генерируем config.xml
            ConfigBuilder.Generate(
                version,
                word,
                excel,
                ppt,
                outlook,
                access,
                publisher,
                lync,
                onedrive
            );

            string setup = ResourceExtractor.ExtractSetup();

            if (!File.Exists(setup))
            {
                MessageBox.Show("setup.exe not found");
                return;
            }

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = setup,
                Arguments = "/configure config.xml",
                WorkingDirectory = Application.StartupPath,
                UseShellExecute = false
            };

            Process.Start(psi);
        }
    }
}