using System;

using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace activador
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private async void materialButton1_Click(object sender, EventArgs e)
        {
            // Supongamos que tienes 8 tareas en total
            materialProgressBar1.Maximum = 5;
            materialProgressBar1.Value = 0;

            string officeDirectory = GetOfficeDirectory();

            if (string.IsNullOrEmpty(officeDirectory))
            {
                terminalRichTextBox.Text += "No se encontró la carpeta de Office.\n";
                return;
            }
            materialProgressBar1.Value++;
            await Task.Delay(900);

            string officeVersion = GetOfficeVersion(officeDirectory);
            terminalRichTextBox.Text += $"Versión de Office detectada: {officeVersion}\n";
            materialProgressBar1.Value++;

            string oldKey = GetOldKey(officeDirectory);
            if (!string.IsNullOrEmpty(oldKey))
            {
                terminalRichTextBox.Text += $"Los últimos 5 caracteres de la clave de producto actual: {oldKey}\n";
            }
            else
            {
                terminalRichTextBox.Text += "No se encontró la clave de producto actual.\n";
            }
            materialProgressBar1.Value++;

            InstallLicenses(officeDirectory, officeVersion);
            materialProgressBar1.Value++;

            ActivateOffice(officeDirectory, oldKey);
            materialProgressBar1.Value++;
            await Task.Delay(900);

            // Muestra el mensaje de finalización en color rojo
            terminalRichTextBox.SelectionStart = terminalRichTextBox.Text.Length;
            terminalRichTextBox.SelectionColor = Color.LightGreen;
            terminalRichTextBox.AppendText("Activación de Office completada.\n");
            terminalRichTextBox.SelectionColor = terminalRichTextBox.ForeColor;

            // Cuando hayas terminado todas las tareas, restablece la barra de progreso a cero
            materialProgressBar1.Value = 0;
        }
        private string GetOfficeDirectory()
        {
            string directory = Path.Combine(Environment.GetEnvironmentVariable("ProgramW6432"), "Microsoft Office", "Office16");
            terminalRichTextBox.Text = directory;
            return Directory.Exists(directory) ? directory : null;
        }

        private string GetOfficeVersion(string officeDirectory)
        {
            string officeVersion = "";
            string output = ExecuteCommand(officeDirectory, "OSPP.VBS /dstatus");
            foreach (string line in output.Split('\n'))
            {
                if (line.Contains("LICENSE NAME"))
                {
                    if (line.Contains("Office16"))
                    {
                        officeVersion = "Office 2016";
                    }
                    else if (line.Contains("Office19"))
                    {
                        officeVersion = "Office 2019";
                    }
                    else if (line.Contains("Office21"))
                    {
                        officeVersion = "Office 2021 LTSC Professional Plus";
                    }
                    break;
                }
            }
            return officeVersion;
        }

        private string GetOldKey(string officeDirectory)
        {
            string output = ExecuteCommand(officeDirectory, "OSPP.VBS /dstatus");
            return output.Split('\n')
                         .FirstOrDefault(line => line.Contains("Last 5 characters of installed product key"))
                         ?.Split(':')
                         .LastOrDefault()
                         ?.Trim();
        }


        private void InstallLicenses(string officeDirectory, string officeVersion)
        {
            string licenseFilePattern;

            switch (officeVersion)
            {
                case "Office 2016":
                    licenseFilePattern = @"ProPlus2016VL_KMS*.xrm-ms";
                    break;
                case "Office 2019":
                    licenseFilePattern = @"ProPlus2019VL_KMS*.xrm-ms";
                    break;
                case "Office 2021 LTSC Professional Plus":
                    licenseFilePattern = @"ProPlus2021VL_KMS*.xrm-ms";
                    break;
                default:
                    return; // No hacemos nada si no reconocemos la versión de Office
            }

            string licenseDirPath = Path.GetFullPath(@"..\root\Licenses16\");
            string[] licenseFiles = Directory.GetFiles(licenseDirPath, licenseFilePattern);

            foreach (string licenseFile in licenseFiles)
            {
                ExecuteCommand(officeDirectory, $"ospp.vbs /inslic:\"{licenseFile}\"");
            }
        }


        private void ActivateOffice(string officeDirectory, string oldKey)
        {
            string port = "1688";
            string newKey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH";
            string kmsHost = "s8.uk.to";

            ExecuteCommand(officeDirectory, $"ospp.vbs /setprt:{port}");
            ExecuteCommand(officeDirectory, $"ospp.vbs /unpkey:{oldKey}");
            ExecuteCommand(officeDirectory, $"ospp.vbs /inpkey:{newKey}");
            ExecuteCommand(officeDirectory, $"ospp.vbs /sethst:{kmsHost}");
            ExecuteCommand(officeDirectory, "ospp.vbs /act");
        }

        private string ExecuteCommand(string directory, string command)
        {

            // Si se proporciona un directorio, cambiar a ese directorio
            if (!string.IsNullOrEmpty(directory))
            {
                try
                {
                    Directory.SetCurrentDirectory(directory);
                }
                catch (Exception e)
                {
                    return $"Error al cambiar al directorio {directory}: {e.Message}";
                }
            }

            bool isVbsCommand = command.ToLower().Contains(".vbs");

            string executable;
            string arguments;

            // Si es un comando VBS, usamos cscript y el script VBS está en el directorio System32
            if (isVbsCommand)
            {
                string system32Directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "slmgr.vbs");
                executable = "cscript";
                arguments = $"//nologo \"{system32Directory}\" {command.Replace("slmgr.vbs", "")}";
            }
            else
            {
                executable = "cmd";
                arguments = $"/c {command}";
            }

            ProcessStartInfo processStartInfo = new ProcessStartInfo(executable, arguments)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            try
            {
                using (Process process = Process.Start(processStartInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    string errorOutput = process.StandardError.ReadToEnd();
                    // Muestra el comando y la salida en terminalRichTextBox
                    terminalRichTextBox.AppendText($"\nComando ejecutado: cmd.exe /c {command}\n");
                    /*terminalRichTextBox.AppendText($"Salida:\n{output}\n");*/

                    // Si hay salida de error, muéstrala también
                    if (!string.IsNullOrEmpty(errorOutput))
                    {
                        terminalRichTextBox.AppendText($"Salida de error:\n{errorOutput}\n");
                    }

                    // Desplaza el scroll hasta el final
                    terminalRichTextBox.SelectionStart = terminalRichTextBox.Text.Length;
                    terminalRichTextBox.ScrollToCaret();

                    return output;
                }
            }
            catch (Exception e)
            {
                return $"Error al ejecutar el comando {command}: {e.Message}";
            }
        }
        private async void materialButton2_Click(object sender, EventArgs e)
        {
            materialProgressBar1.Minimum = 0;
            materialProgressBar1.Maximum = 3;
            materialProgressBar1.Value = 0;
            materialProgressBar1.Step = 1;
            string windowsEdition;
            try
            {
                windowsEdition = GetWindowsEdition();
            }
            catch (Exception ex)
            {
                terminalRichTextBox.AppendText($"Error al obtener la edición de Windows: {ex.Message}\n");
                return;
            }

            terminalRichTextBox.AppendText($"Edición de Windows: {windowsEdition}\n");

            var keys = new Dictionary<string, List<string>>
{
    {"Microsoft Windows 10 Pro", new List<string>{"VK7JG-NPHTM-C97JM-9MPGT-3V66T", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J", "W269N-WFGWX-YVC9B-4J6C9-T83GX"}},
    {"Microsoft Windows 10 Pro Education", new List<string>{"6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"}},
    {"Microsoft Windows 10 Pro Education N", new List<string>{"YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"}},
    {"Microsoft Windows 10 Pro N", new List<string>{"MH37W-N47XK-V7XM9-C7227-GCQG9", "9FNHH-K3HBT-3W4TD-6383H-6XYWF"}},
    {"Microsoft Windows 10 Home", new List<string>{"TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"}},
    {"Microsoft Windows 10 Home Single Language", new List<string>{"7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"}},
    {"Microsoft Windows 10 Education", new List<string>{"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"}},
    {"Microsoft Windows 10 Education N", new List<string>{"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"}},
    {"Microsoft Windows 10 Enterprise", new List<string>{"NPPR9-FWDCX-D2C8J-H872K-2YT43"}},
    {"Microsoft Windows 10 Enterprise G", new List<string>{"YYVX9-NTFWV-6MDM3-9PT4T-4M68B"}},
    {"Microsoft Windows 10 Enterprise G N", new List<string>{"44RPN-FTY23-9VTTB-MP9BX-T84FV"}},
    {"Microsoft Windows 10 Enterprise N", new List<string>{"DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"}},
    {"Microsoft Windows 11 Home", new List<string>{"TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"}},
    {"Microsoft Windows 11 Home Single Language", new List<string>{"7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"}},
    {"Microsoft Windows 11 Pro", new List<string>{"W269N-WFGWX-YVC9B-4J6C9-T83GX"}},
    {"Microsoft Windows 11 Pro N", new List<string>{"MH37W-N47XK-V7XM9-C7227-GCQG9"}},
    {"Microsoft Windows 11 Education", new List<string>{"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"}},
    {"Microsoft Windows 11 Education N", new List<string>{"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"}},
    {"Microsoft Windows 11 Enterprise", new List<string>{"NPPR9-FWDCX-D2C8J-H872K-2YT43"}},
        {"Microsoft Windows 11 Enterprise N", new List<string>{"DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"}},
    {"Microsoft Windows 11 Enterprise G", new List<string>{"YYVX9-NTFWV-6MDM3-9PT4T-4M68B"}},
    {"Microsoft Windows 11 Enterprise G N", new List<string>{"44RPN-FTY23-9VTTB-MP9BX-T84FV"}},
    {"Microsoft Windows 11 Workstations", new List<string>{"NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"}},
    {"Microsoft Windows 11 Ultimate", new List<string>{"Q269N-WFGWX-YVC9B-4J6C9-T83GX"}}
};



            if (keys.ContainsKey(windowsEdition))
            {
                string selectedKey = keys[windowsEdition][0];
                terminalRichTextBox.AppendText($"Clave seleccionada para {windowsEdition}: {selectedKey}\n");

                ExecuteCommand("", $"slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX");
                string[] kmsHosts = { "kms.digiboy.ir", "kms.msguides.com", "kms7.MSGuides.com:1688" };
                bool activationSuccessful = false;

                materialProgressBar1.PerformStep(); // Esto incrementará el valor del ProgressBar en 1
                await Task.Delay(900);
                foreach (string kmsHost in kmsHosts)
                {
                    ExecuteCommand("", $"slmgr.vbs /skms {kmsHost}");
                    string activationResult =  ExecuteCommand("", "slmgr.vbs /ato");

                    if (activationResult.ToLower().Contains("successfully"))
                    {
                        terminalRichTextBox.Invoke((MethodInvoker)async delegate {
                            terminalRichTextBox.SelectionStart = terminalRichTextBox.Text.Length;
                            terminalRichTextBox.SelectionColor = Color.LightGreen;
                            terminalRichTextBox.AppendText($"Activación de Windows exitosa");
                            terminalRichTextBox.SelectionColor = terminalRichTextBox.ForeColor;
                        });
                        activationSuccessful = true;
                        break;
                    }
                    else
                    {
                        terminalRichTextBox.Invoke((MethodInvoker)async delegate {
                            terminalRichTextBox.AppendText($"Error en la activación con el KMS host: {kmsHost}\n");
                            materialProgressBar1.PerformStep();
                            await Task.Delay(900);
                        });
                    }
                }

                if (!activationSuccessful)
                {
                    terminalRichTextBox.AppendText("No se pudo activar Windows con ninguno de los KMS hosts proporcionados.\n");
                }
            }
            else
            {
                terminalRichTextBox.AppendText("No se encontró una clave para la edición de Windows actual.\n");
            }
            materialProgressBar1.Value = materialProgressBar1.Maximum; // Esto incrementará el valor del ProgressBar en 1
            await Task.Delay(500);
            materialProgressBar1.Value = 0;
        }

        private string GetWindowsEdition()
        {
            try
            {
                var caption = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem")
                                .Get()
                                .Cast<ManagementObject>()
                                .First()["Caption"].ToString();

                return caption;
            }
            catch (Exception e)
            {
                throw new Exception($"Error al obtener la edición de Windows: {e.Message}");
            }
        }
    }
}

