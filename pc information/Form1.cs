using System;
using System.Windows.Forms;
using System.Management;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using System.Text;

namespace pc_information
{
    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectedIndexChanged += tabControl1_TabIndexChanged;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnOS.Enabled = false;
            btnProcess.Enabled = false;
            getInfo();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
           
        }
        public void getInfo()
        {

            try
            {
                // Retrieve operating system information
                ManagementObjectSearcher searcher1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectCollection osCollection = searcher1.Get();
                foreach (ManagementObject os in osCollection)
                {
                    lblOperatingSystem.Text = $"{os["Caption"]} {os["Version"]}";
                }

                // Retrieve BIOS information
                searcher1 = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
                ManagementObjectCollection biosCollection = searcher1.Get();
                foreach (ManagementObject bios in biosCollection)
                {
                    lblBIOSVendor.Text = $"{bios["Version"]}";
                }

                // Retrieve DirectX version from the registry
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\DirectX"))
                {
                    if (key != null)
                    {
                        object directXVersion = key.GetValue("Version");
                        if (directXVersion != null)
                        {
                            lblDirectXVersion.Text = $"{directXVersion.ToString()}";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            String q1 = Environment.MachineName; // Computer Name
            lblComputerName.Text = q1;

            String q2 = Environment.UserName; // User Name 
            lblUserName.Text = Convert.ToString(q2);

            bool q5 = Environment.Is64BitOperatingSystem; // Is you system 64 Bit OS
            lblBitOS.Text = Convert.ToString(q5);
            if (lblBitOS.Text == "True")
            {
                btnOS.Text = "✔";
            }
            else
            {
                btnOS.Text = "✕";
            }

            bool q6 = Environment.Is64BitOperatingSystem; // Is your systemm 64 Bit Process
            lblBitProcess.Text = Convert.ToString(q6);
            if (lblBitProcess.Text == "True")
            {
                btnProcess.Text = "✔";
            }
            else
            {
                btnProcess.Text = "✕";
            }

            String q8 = (Environment.OSVersion.ToString()); // version in word 
            lblVersion.Text = q8;


            String q9 = Environment.OSVersion.Platform.ToString(); // System Platform
            lbloperatingSystemPlatform.Text = q9;


            // Retrieve system model using ManagementObjectSearcher
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject obj in collection)
            {
                string systemModel = obj["Model"].ToString();
                lblSystemModel.Text = systemModel;

                // Memory information
                ManagementObjectSearcher memorySearcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectCollection memoryCollection = memorySearcher.Get();

                foreach (ManagementObject memory in memoryCollection)
                {
                    ulong totalPhysicalMemory = Convert.ToUInt64(memory["TotalVisibleMemorySize"]);
                    ulong freePhysicalMemory = Convert.ToUInt64(memory["FreePhysicalMemory"]);
                    ulong totalVirtualMemory = Convert.ToUInt64(memory["TotalVirtualMemorySize"]);
                    ulong availableVirtualMemory = Convert.ToUInt64(memory["FreeVirtualMemory"]);

                   lblinstalledPhysicalMemory.Text = $"{totalPhysicalMemory / (1024 * 1024)} GB"; // Convert bytes to megabytes
                    lblavailablePhysicalMemory.Text = $"{freePhysicalMemory / (1024 * 1024)} GB";
                    lbltotalVirtualMemory.Text = $"{totalVirtualMemory / (1024 * 1024)} GB";
                    lblavailableVirtualMemory.Text = $"{availableVirtualMemory / (1024 * 1024)} GB";

                }

                // Get information about the page file on disk
                FileInfo pageFileInfo = new FileInfo("C:\\pagefile.sys");

                // Check if the file exists before attempting to get its size
                if (pageFileInfo.Exists)
                {
                    // Get the size of the page file in bytes
                    long pageFileSizeInBytes = pageFileInfo.Length;

                    // Convert the size to gigabytes with two decimal places
                    double pageFileSizeInGB = (double)pageFileSizeInBytes / (1024 * 1024 * 1024);

                    // Display the page file size
                    lblpageFileSpace.Text = $"{pageFileSizeInGB:F2} GB";
                }
                else
                {
                    // Handle the case where the page file does not exist
                    lblpageFileSpace.Text = "Page file not found";
                }


                // Baseboard information
                string baseboardManufacturer = obj["Manufacturer"].ToString();

                // Processor information
                ManagementObjectSearcher processorSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                ManagementObjectCollection processorCollection = processorSearcher.Get();

                foreach (ManagementObject processorObj in processorCollection)
                {
                    // Check if the object and required properties are not null before accessing the values
                    if (processorObj != null &&
                        processorObj["Name"] != null &&
                        processorObj["MaxClockSpeed"] != null &&
                        processorObj["NumberOfCores"] != null &&
                        processorObj["NumberOfLogicalProcessors"] != null)
                    {
                        string processorName = processorObj["Name"].ToString();
                        string maxClockSpeed = (Convert.ToDouble(processorObj["MaxClockSpeed"]) / 1000).ToString("F2") + " GHz";
                        string numberOfCores = processorObj["NumberOfCores"].ToString();
                        string numberOfLogicalProcessors = processorObj["NumberOfLogicalProcessors"].ToString();

                        // Combine the information into a single string
                        processorName = processorName.Substring(0, processorName.Length - 10);
                        string processorInfo = $"Processor {processorName }";

                        // Update your UI element with the processor information
                        // For example:
                        lblprocessor.Text = processorInfo;
                        lblMaxClockSpeed.Text = maxClockSpeed;
                        lblNumberOfLogicalProcessors.Text = numberOfLogicalProcessors;
                        lblNumberOfCores.Text = numberOfCores;
                        break; // Assuming you only want information from the first processor
                    }
                }


                lblBaseBoardManufacture.Text = baseboardManufacturer;

                // Graphics card information
                ManagementObjectSearcher gpuSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");
                ManagementObjectCollection gpuCollection = gpuSearcher.Get();

                

            }

            // Retrieve DirectX version from the registry
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("HARDWARE\\DESCRIPTION\\System\\BIOS"))
            {
                if (key != null)
                {
                    object BIOSReleaseDate = key.GetValue("BIOSReleaseDate");
                    if (BIOSReleaseDate != null)
                    {
                        lblBIOSReleaseDate.Text = $"{BIOSReleaseDate.ToString()}";
                    }
                    object BIOSVendor = key.GetValue("BIOSVendor");
                    if (BIOSVendor != null)
                    {
                        lblBIOSVendor.Text = $"{BIOSVendor.ToString()}";
                    }
                    object BIOSVersion = key.GetValue("BIOSVersion");
                    if (BIOSVersion != null)
                    {
                        lblBIOSVersion.Text = $"{BIOSVersion.ToString()}";
                    }
                }
            }
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0"))
            {
                if (key != null)
                {
                    object Identifier = key.GetValue("Identifier");
                    if (Identifier != null)
                    {
                        lblProcessorFamily.Text = $"{Identifier.ToString()}";
                    }
                   
                }
            }
            try
            {
                // PowerShell command to get GPU information
                string powerShellCommand = "Get-CimInstance win32_VideoController | Select-Object Name, DeviceID, VideoProcessor,AdapterDACType, AdapterCompatibility, VideoModeDescription, MaxRefreshRate, CurrentRefreshRate, DriverVersion,Status";

                // Create a process to run PowerShell
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = $"-NoProfile -ExecutionPolicy unrestricted -Command \"{powerShellCommand}\""
                };

                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();

                    // Read the output of the PowerShell command
                    string output = process.StandardOutput.ReadToEnd().Trim();

                    process.WaitForExit();

                    // Split the output into lines
                    string[] gpuLines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    // Display GPU information in separate labels
                    if (gpuLines.Length > 0)
                    {
                        lblGraphicsCardName.Text = $"{gpuLines[0]}";
                        if (gpuLines.Length > 1)
                            lblDeviceID.Text = $"{gpuLines[1]}";

                        if (gpuLines.Length > 2)
                            lblVideoProcessor.Text = $"{gpuLines[2]}";
                        if (gpuLines.Length > 3)
                            lblDacType.Text = $"{gpuLines[3]}";
                        if (gpuLines.Length > 4)
                            lblAdapterCompatibility.Text = $"{gpuLines[4]}";
                        if (gpuLines.Length > 5)
                            lblVideoModeDescription.Text = $"{gpuLines[5]}";

                        if (gpuLines.Length > 6)
                            lblMaxRefreshRate.Text = $"{gpuLines[6]}";

                        if (gpuLines.Length > 7)
                            lblCurrentRefreshRate.Text = $"{gpuLines[7]}";

                        if (gpuLines.Length > 8)
                            lblDriverVersion.Text = $"{gpuLines[8]}";
                        if (gpuLines.Length > 9)
                            lblStatus.Text = $"{gpuLines[9]}";
                        // Add similar code for other properties you want to display

                        // If you have more labels, you can continue the pattern
                    }
                    else
                    {
                        lblGraphicsCardName.Text = "Graphics Card information not available";
                    }
                    // Display GPU information in separate labels
                    if (gpuLines.Length > 10)  // Start reading from the information of the second GPU
                    {
                        lblGraphicsCardName2.Text = $"{gpuLines[10]}";

                        if (gpuLines.Length > 11)
                            lblDeviceID2.Text = $"{gpuLines[11]}";

                        if (gpuLines.Length > 15)
                           lblVideoModeDescription2.Text = $"{gpuLines[15]}";

                        if (gpuLines.Length > 16)
                            lblMaxRefreshRate2.Text = $"{gpuLines[16]}";

                        if (gpuLines.Length > 17)
                            lblCurrentRefreshRate2.Text = $"{gpuLines[17]}";
                        if (gpuLines.Length > 18)
                            lblStatus2.Text = $"{gpuLines[18]}";
                        if (gpuLines.Length > 19)
                            lblStatus2.Text = $"{gpuLines[19]}";

                        // If you have more labels, add similar code for other properties you want to display
                    }
                    else
                    {
                        lblGraphicsCardName2.Text = "Graphics Card information not available";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                // PowerShell command to get sound information
                string powerShellCommand = "Get-WmiObject Win32_SoundDevice | Select-Object Manufacturer, Name, Status | Format-Table | Out-String -Width 120";

                // Create a process to run PowerShell
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = $"-NoProfile -ExecutionPolicy unrestricted -Command \"{powerShellCommand}\""
                };

                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();

                    // Read the output of the PowerShell command
                    string output = process.StandardOutput.ReadToEnd().Trim();

                    process.WaitForExit();

                    // Split the output into lines
                    string[] soundLines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    // Display sound information in separate labels
                    if (soundLines.Length >= 5)  // Check if there are at least 3 lines (header + 3 properties)
                    {
                        // Manufacturer
                        lblManufacturer.Text = $"{soundLines[2]}";

                        // Name
                        lblName.Text = $"{soundLines[3]}";

                        // Status
                        lblStatusSound.Text = $"{soundLines[4]}";
                    }
                    else
                    {
                        lblManufacturer.Text = "Manufacturer information not available";
                        lblName.Text = "Name information not available";
                        lblStatusSound.Text = "Status information not available";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                // PowerShell command to get input devices information
                string powerShellCommand = "Get-WmiObject Win32_PointingDevice | Select-Object DeviceID, Manufacturer, Description, Status | Format-Table | Out-String -Width 120";

                // Create a process to run PowerShell
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = $"-NoProfile -ExecutionPolicy unrestricted -Command \"{powerShellCommand}\""
                };

                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();

                    // Read the output of the PowerShell command
                    string output = process.StandardOutput.ReadToEnd().Trim();

                    process.WaitForExit();

                    // Display input devices information in lblInputDevices
                    if (!string.IsNullOrEmpty(output))
                    {
                        lblInputDevices1.Text = output;
                    }
                    else
                    {
                        lblInputDevices1.Text = "Input devices information not available";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                // PowerShell command to get input devices information
                string keyboardCommand = "Get-WmiObject Win32_PnPEntity | Where-Object { $_.Caption -like '*keyboard*' } | Select-Object DeviceID, Caption, Status | Format-Table | Out-String -Width 120";
                string wirelessButtonCommand = "Get-WmiObject Win32_PnPEntity | Where-Object { $_.Caption -like '*wireless button*' } | Select-Object DeviceID, Caption, Status | Format-Table | Out-String -Width 120";

                string keyboardOutput = ExecutePowerShellCommand(keyboardCommand);
                string wirelessButtonOutput = ExecutePowerShellCommand(wirelessButtonCommand);

                // Display input devices information in lblInputDevices2 and lblInputDevices3
                if (!string.IsNullOrEmpty(keyboardOutput))
                {
                    lblInputDevices2.Text = keyboardOutput;
                }
                else
                {
                    lblInputDevices2.Text = "Keyboard information not available";
                }

                if (!string.IsNullOrEmpty(wirelessButtonOutput))
                {
                    lblInputDevices3.Text = wirelessButtonOutput;
                }
                else
                {
                    lblInputDevices3.Text = "Wireless button information not available";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Helper method to execute PowerShell command
           
        }
        private string ExecutePowerShellCommand(string command)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "powershell.exe";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.Arguments = $"-NoProfile -ExecutionPolicy unrestricted -Command \"{command}\"";

                process.Start();

                // Read the output of the PowerShell command
                string output = process.StandardOutput.ReadToEnd().Trim();

                process.WaitForExit();

                return output;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
                Application.Exit();
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string aboutMessage = "PC Information 2024\n" +
                "Version 2.0\n" +
                "© 2024 Ahadu Tech\n" +
                "Initial Creator Nahu Senay\n" +
                "Enhanced by NattyXO";

            MessageBox.Show(aboutMessage, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void picGithub_Click(object sender, EventArgs e)
        {
            string githublink = "https://github.com/NattyXO";
            Process.Start(githublink);
        }

        private void menuStrip1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void btnNextpage_Click(object sender, EventArgs e)
        {
            int nextTabIndex = tabControl1.SelectedIndex + 1;

            if (nextTabIndex < tabControl1.TabCount)
            {
                tabControl1.SelectedIndex = nextTabIndex;
                UpdateNextPageButtonVisibility();
            }
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            UpdateNextPageButtonVisibility();
        }
        private void UpdateNextPageButtonVisibility()
        {
            btnNextpage.Enabled = (tabControl1.SelectedIndex < tabControl1.TabCount - 1);
        }

        private void btnSaveInfo_Click(object sender, EventArgs e)
        {
            // Create a StringBuilder to store the information
            StringBuilder pcInfoStringBuilder = new StringBuilder();

            // Define the labels for each section
            List<(string, Label)> systemLabels = new List<(string, Label)>
    {
        ("Computer Name: ", lblComputerName),
        ("User Name: ", lblUserName),
        ("OS Version: ", lblVersion),
        ("Operating System: ", lblOperatingSystem),
        ("OS Platform: ", lbloperatingSystemPlatform),
        ("System Model: ", lblSystemModel),
        ("Base Board Manufacture: ", lblBaseBoardManufacture),
        ("DirectX Version: ", lblDirectXVersion),
        ("Processor: ", lblprocessor),
        ("Processor Family: ", lblProcessorFamily),
        ("Max Clock Speed: ", lblMaxClockSpeed),
        ("Number Of Cores: ", lblNumberOfCores),
        ("Number Of Logical Processors: ", lblNumberOfLogicalProcessors),
        ("Installed Physical Memory: ", lblinstalledPhysicalMemory),
        ("Available Physical Memory: ", lblavailablePhysicalMemory),
        ("Total Virtual Memory: ", lbltotalVirtualMemory),
        ("Available Virtual Memory: ", lblavailableVirtualMemory),
        ("Page File Space: ", lblpageFileSpace),
        ("BIOS Vendor: ", lblBIOSVendor),
        ("BIOS Version: ", lblBIOSVersion),
        ("BIOS Release Date: ", lblBIOSReleaseDate),
        ("64-Bit OS: ", lblBitOS),
        ("64-Bit Process: ", lblBitProcess)
    };

            List<(string, Label)> displayLabels = new List<(string, Label)>
    {
        ("Graphics Card Name 1\n", lblGraphicsCardName),
        ("", lblVideoModeDescription),
        ("", lblVideoProcessor),
        ("", lblDeviceID),
        ("", lblAdapterCompatibility),
        ("", lblMaxRefreshRate),
        ("", lblCurrentRefreshRate),
        ("", lblDriverVersion),
        ("", lblStatus),
        ("\nGraphics Card Name 2 \n", lblGraphicsCardName2),
        ("", lblDeviceID2),
        ("", lblMaxRefreshRate2),
        ("", lblVideoModeDescription2),
        ("", lblStatus2)
    };

            List<(string, Label)> soundLabels = new List<(string, Label)>
    {
        ("", lblManufacturer),
        ("", lblName),
        ("", lblStatusSound)
    };

            List<(string, Label)> inputLabels = new List<(string, Label)>
    {
        ("Input Devices 1: \n", lblInputDevices1),
        ("\nInput Devices 2: \n", lblInputDevices2),
        ("\nInput Devices 3: \n", lblInputDevices3),
        ("\n\n\n", lblSystemAbout)
    };

            // Add section headers and labels to the StringBuilder
            AppendSection(pcInfoStringBuilder, "--- System ---", systemLabels);
            AppendSection(pcInfoStringBuilder, "--- Display ---", displayLabels);
            AppendSection(pcInfoStringBuilder, "--- Sound ---", soundLabels);
            AppendSection(pcInfoStringBuilder, "--- Input ---", inputLabels);

            // Save the information to a text file on the desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "PC_Info.txt");
            SaveToFile(pcInfoStringBuilder.ToString(), filePath);
        }

        private void AppendSection(StringBuilder stringBuilder, string sectionHeader, List<(string, Label)> labels)
        {
            // Append the section header
            stringBuilder.AppendLine(sectionHeader);

            // Iterate through labels and append their names and values
            foreach ((string labelName, Label label) in labels)
            {
                stringBuilder.AppendLine($"{labelName}{label.Text}");
            }

            // Add a separator between sections
            stringBuilder.AppendLine();
        }

        private void SaveToFile(string content, string filePath)
        {
            // Save content to the specified file path
            try
            {
                // Write the content to the file
                File.WriteAllText(filePath, content);

                MessageBox.Show($"PC information saved to {filePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       
    }
}
