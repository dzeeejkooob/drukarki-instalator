using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;

namespace drukarki
{
    public partial class Form1 : Form
    {
        private List<(string printerName, string ipAddress)> printers = new List<(string, string)>
        {
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP"),
            ("Nazwa drukarki", "adres IP")
        };

        public Form1()
        {
            InitializeComponent();
            InitializePrinterList();

            if (!IsRunningAsAdministrator())
            {
                MessageBox.Show("Uwaga! Aplikacja nie jest uruchomiona jako administrator. Dodawanie drukarek może nie zadziałać poprawnie.", "Brak uprawnień", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
        }

        private void InitializePrinterList()
        {
            foreach (var printer in printers)
            {
                listBox1.Items.Add(printer.printerName);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    listBox1.SetSelected(i, true);
                }
            }
            else
            {
                listBox1.ClearSelected();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var originalColor = button1.BackColor;
            button1.BackColor = Color.Orange;
            button1.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            StringBuilder log = new StringBuilder();
            List<(string printerName, string ipAddress)> selectedPrinters = new List<(string, string)>();

            if (checkBox1.Checked)
            {
                selectedPrinters.AddRange(printers);
            }
            else
            {
                if (listBox1.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Nie wybrano żadnej drukarki.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                foreach (var selectedPrinter in listBox1.SelectedItems)
                {
                    var printer = printers.Find(p => p.printerName == selectedPrinter.ToString());
                    if (printer != default)
                    {
                        selectedPrinters.Add(printer);
                    }
                }
            }

            foreach (var printer in selectedPrinters)
            {
                string result = AddPrinter(printer.printerName, printer.ipAddress);
                log.AppendLine(result);
            }

            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string logPath = Path.Combine(documentsPath, "logi_drukarki.txt");
            File.AppendAllText(logPath, $"{DateTime.Now:G}\n{log}\n");

            button1.BackColor = originalColor;
            button1.Enabled = true;
            Cursor.Current = Cursors.Default;

            MessageBox.Show(log.ToString(), "Wynik", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string AddPrinter(string printerName, string printerIP)
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string driverPath = Path.Combine(appDirectory, "KOAX7J__.inf");

            if (!File.Exists(driverPath))
            {
                return $"{printerName}: Nie znaleziono pliku sterownika w: {driverPath}";
            }

            var installDriverStartInfo = new ProcessStartInfo
            {
                FileName = @"C:\Windows\Sysnative\pnputil.exe",
                Arguments = $"/add-driver \"{driverPath}\" /install",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var process = Process.Start(installDriverStartInfo))
            {
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(errorOutput))
                {
                    return $"{printerName}: Błąd podczas instalacji sterownika .inf: {errorOutput}";
                }
            }

            string driverBaseName = "KONICA MINOLTA Universal";
            var getDriverStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -Command \"Get-PrinterDriver | Where-Object {{$_.Name -like '{driverBaseName}*'}} | Select-Object -First 1 -ExpandProperty Name\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            string resolvedDriverName;
            using (var process = Process.Start(getDriverStartInfo))
            {
                resolvedDriverName = process.StandardOutput.ReadToEnd().Trim();
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (string.IsNullOrWhiteSpace(resolvedDriverName))
                {
                    return $"{printerName}: Nie znaleziono pasującego sterownika drukarki (szukano: '{driverBaseName}*')";
                }
            }

            var addPortStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -Command \"if (-not (Get-PrinterPort -Name 'TCPIP_{printerName}' -ErrorAction SilentlyContinue)) " +
                            $"{{ Add-PrinterPort -Name 'TCPIP_{printerName}' -PrinterHostAddress {printerIP} }}\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var process = Process.Start(addPortStartInfo))
            {
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(errorOutput))
                {
                    return $"{printerName}: Błąd podczas dodawania portu: {errorOutput}";
                }
            }

            var addPrinterStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -Command \"if (-not (Get-Printer -Name '{printerName}' -ErrorAction SilentlyContinue)) " +
                    $"{{ Add-Printer -Name '{printerName}' -PortName 'TCPIP_{printerName}' -DriverName '{resolvedDriverName}' }}\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var process = Process.Start(addPrinterStartInfo))
            {
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(errorOutput))
                {
                    return $"{printerName}: Błąd podczas dodawania drukarki: {errorOutput}";
                }
            }

            return $"{printerName}: Drukarka dodana pomyślnie!";
        }

        private bool IsRunningAsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private void pomocToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string instrukcja = "Instrukcja użytkowania:\n\n" +
                                "1. Wybierz z listy drukarki, które chcesz dodać.\n" +
                                "2. Kliknij przycisk 'Dodaj'.\n" +
                                "3. Poczekaj, aż program zakończy instalację sterowników oraz konfigurację drukarki.\n" +
                                "4. Po zakończeniu instalacji drukarka będzie gotowa do użycia.\n\n" +
                                "W przypadku pytań lub problemów, prosimy o kontakt: wsparcie@";
            MessageBox.Show(instrukcja, "Pomoc", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string info = "Nazwa aplikacji: Instalator drukarek IBE\n" +
                  "Wersja: 1.0.0\n" +
                  "Autor: Jakub Wojciechowski\n" +
                  "Instytucja: Instytut Badań Edukacyjnych\n\n" +
                  "Aplikacja umożliwia szybkie i automatyczne dodawanie drukarek wraz z odpowiednimi sterownikami.\n\n" +
                  "Wszelkie pytania i zgłoszenia błędów prosimy kierować na adres: wsparcie@";
            MessageBox.Show(info, "Informacje o aplikacji", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Zareaguj na zmianę wybranego elementu w ListBox
            MessageBox.Show($"Wybrano: {listBox1.SelectedItem}");
        }
    }
}
