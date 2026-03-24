using System.IO;
using System.Text;
using System.Windows.Forms;

namespace OfficeInstaller
{
    public static class ConfigBuilder
    {
        public static void Generate(
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
            string product = "O365ProPlusRetail";
            string channel = "Current";

            switch (version)
            {
                case "2019":
                    product = "ProPlus2019Volume";
                    channel = "PerpetualVL2019";
                    break;

                case "2021":
                    product = "ProPlus2021Volume";
                    channel = "PerpetualVL2021";
                    break;

                case "2024":
                    product = "ProPlus2024Volume";
                    channel = "PerpetualVL2021";
                    break;
            }

            StringBuilder xml = new StringBuilder();

            xml.AppendLine("<Configuration>");
            xml.AppendLine("  <Remove All=\"TRUE\" />");
            xml.AppendLine($"  <Add OfficeClientEdition=\"64\" Channel=\"{channel}\">");
            xml.AppendLine($"    <Product ID=\"{product}\">");
            xml.AppendLine("      <Language ID=\"MatchOS\" />");

            if (!word)
                xml.AppendLine("      <ExcludeApp ID=\"Word\" />");

            if (!excel)
                xml.AppendLine("      <ExcludeApp ID=\"Excel\" />");

            if (!ppt)
                xml.AppendLine("      <ExcludeApp ID=\"PowerPoint\" />");

            if (!outlook)
                xml.AppendLine("      <ExcludeApp ID=\"Outlook\" />");

            if (!access)
                xml.AppendLine("      <ExcludeApp ID=\"Access\" />");

            if (!publisher)
                xml.AppendLine("      <ExcludeApp ID=\"Publisher\" />");

            if (!lync)
                xml.AppendLine("      <ExcludeApp ID=\"Lync\" />");

            if (!onedrive)
                xml.AppendLine("      <ExcludeApp ID=\"OneDrive\" />");

            xml.AppendLine("    </Product>");
            xml.AppendLine("  </Add>");
            xml.AppendLine("  <Display Level=\"Full\" AcceptEULA=\"TRUE\" />");
            xml.AppendLine("  <Updates Enabled=\"TRUE\" Channel=\"" + channel + "\" />");
            xml.AppendLine("</Configuration>");

            string path = Path.Combine(Application.StartupPath, "config.xml");

            File.WriteAllText(path, xml.ToString(), Encoding.UTF8);
        }
    }
}