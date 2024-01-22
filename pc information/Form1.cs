using System;
using System.Windows.Forms;
using System.Management;
using System.IO;
using System.Diagnostics;

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
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnOS.Enabled = false;
            btnProcess.Enabled = false;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            txtNumberOfCores.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            btnProcess.Text = "?";
            btnOS.Text = "?";
            txtavailablePhysicalMemory.Text = "";
            txtavailableVirtualMemory.Text = "";
            txtbaseboardManufacturer.Text = "";
            txtBiosMode.Text = "";
            txtBiosVersion.Text = "";
            txtgraphicsCardMemory.Text = "";
            txtgraphicsCardName.Text = "";
            txtinstalledPhysicalMemory.Text = "";
            txtpageFileSpace.Text = "";
            txtprocessor.Text = "";
            txtSystemModel.Text = "";
            txttotalVirtualMemory.Text = "";
            txtMaxClockSpeed.Text = "";
            txtNumberOfCores.Text = "";
            txtNumberOfLogicalProcessors.Text = "";
        }

        private void btnGetInfo_Click(object sender, EventArgs e)
        {
            textBox4.Enabled = true;
            textBox5.Enabled = true;


            String q1 = Environment.MachineName; // Computer Name
            textBox1.Text = q1;

            String q2 = Environment.UserName; // User Name 
            textBox2.Text = Convert.ToString(q2);

            bool q5 = Environment.Is64BitOperatingSystem; // Is you system 64 Bit OS
            textBox4.Text = Convert.ToString(q5);
            if (textBox4.Text == "True")
            {
                btnOS.Text = "✔";
            }else
            {
                btnOS.Text = "✕";
            }

            bool q6 = Environment.Is64BitOperatingSystem; // Is your systemm 64 Bit Process
            textBox5.Text = Convert.ToString(q6);
            if (textBox5.Text == "True")
            {
                btnProcess.Text = "✔";
            }
            else
            {
                btnProcess.Text = "✕";
            }

            String q8 = (Environment.OSVersion.ToString()); // version in word 
            textBox6.Text = q8;


            String q9 = Environment.OSVersion.Platform.ToString(); // System Platform
            textBox7.Text = q9;


            string q10 = Environment.OSVersion.Platform.ToString(); // System Platform
            textBox8.Text = q10;

            // Retrieve system model using ManagementObjectSearcher
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject obj in collection)
            {
                string systemModel = obj["Model"].ToString();
                txtSystemModel.Text = systemModel;

                try
                {
                    // BIOS information
                    PropertyData biosVersionProperty = obj.Properties["BIOSVersion"];
                    PropertyData biosModeProperty = obj.Properties["BIOSMode"];

                    string biosVersion = "N/A";
                    string biosMode = "N/A";

                    // Check if the properties exist before trying to access them
                    if (biosVersionProperty != null && biosModeProperty != null)
                    {
                        // Check if the values are not null and assign them to the strings
                        if (biosVersionProperty.Value != null)
                            biosVersion = biosVersionProperty.Value.ToString();

                        if (biosModeProperty.Value != null)
                            biosMode = biosModeProperty.Value.ToString();
                    }

                    // Update your UI elements with biosVersion and biosMode as needed
                    // For example:
                    txtBiosVersion.Text = biosVersion;
                    txtBiosMode.Text = biosMode;
                }
                catch (Exception ex)
                {
                    // Log or handle the exception here
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    // You might want to log the exception details to a file or another logging mechanism.
                    // Optionally, you can show an error message to the user or take other appropriate actions.
                }

                // Memory information
                ManagementObjectSearcher memorySearcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectCollection memoryCollection = memorySearcher.Get();

                foreach (ManagementObject memory in memoryCollection)
                {
                    ulong totalPhysicalMemory = Convert.ToUInt64(memory["TotalVisibleMemorySize"]);
                    ulong freePhysicalMemory = Convert.ToUInt64(memory["FreePhysicalMemory"]);
                    ulong totalVirtualMemory = Convert.ToUInt64(memory["TotalVirtualMemorySize"]);
                    ulong availableVirtualMemory = Convert.ToUInt64(memory["FreeVirtualMemory"]);
                    
                    txtinstalledPhysicalMemory.Text = $"{totalPhysicalMemory / (1024 * 1024)} GB"; // Convert bytes to megabytes
                    txtavailablePhysicalMemory.Text = $"{freePhysicalMemory / (1024 * 1024)} GB";
                    txttotalVirtualMemory.Text = $"{totalVirtualMemory / (1024 * 1024)} GB";
                    txtavailableVirtualMemory.Text = $"{availableVirtualMemory / (1024 * 1024)} GB";
                   
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
                    txtpageFileSpace.Text = $"{pageFileSizeInGB:F2} GB";
                }
                else
                {
                    // Handle the case where the page file does not exist
                    txtpageFileSpace.Text = "Page file not found";
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
                        txtprocessor.Text = processorInfo;
                        txtMaxClockSpeed.Text = maxClockSpeed;
                        txtNumberOfLogicalProcessors.Text = numberOfLogicalProcessors;
                        txtNumberOfCores.Text = numberOfCores;
                        break; // Assuming you only want information from the first processor
                    }
                }


                txtbaseboardManufacturer.Text = baseboardManufacturer;

                // Graphics card information
                ManagementObjectSearcher gpuSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");
                ManagementObjectCollection gpuCollection = gpuSearcher.Get();

                foreach (ManagementObject gpu in gpuCollection)
                {
                    // Check if the object is not null before accessing its properties
                    if (gpu != null)
                    {
                        object adapterRAMObj = gpu["AdapterRAM"];
                        object captionObj = gpu["Caption"];

                        // Check if the properties are not null before converting and using them
                        if (adapterRAMObj != null && captionObj != null)
                        {
                            string graphicsCardMemoryStr = adapterRAMObj.ToString();
                            string graphicsCardName = captionObj.ToString();

                            // Try to parse the values, and if successful, update UI elements
                            if (ulong.TryParse(graphicsCardMemoryStr, out ulong graphicsCardMemory))
                            {
                                txtgraphicsCardMemory.Text = $"{graphicsCardMemory / (1024 * 1024)} MB";
                            }
                            else
                            {
                                // Handle the case where parsing fails
                                txtgraphicsCardMemory.Text = "N/A";
                            }

                            txtgraphicsCardName.Text = graphicsCardName;
                        }
                        else
                        {
                            // Handle the case where properties are null
                            txtgraphicsCardMemory.Text = "N/A";
                            txtgraphicsCardName.Text = "N/A";
                        }
                    }
                }

            }
            

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            const string message =
                "Would you like to close?";
            const string caption = "Information";
            var result = MessageBox.Show(message, caption,
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);

            //if the no button was pressed...
            if (result == DialogResult.Yes)
            {
                //cancel the Closur of the form.
                Application.Exit();
            }
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string aboutMessage = "PC Information 2024\n" +
                "Version 1.0\n" +
                "© 2024 Ahadu Tech\n"+
                "Initial Creator Nahu Senay\n"+
                "Enhanced by NattyXO";

            MessageBox.Show(aboutMessage, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void githubToolStripMenuItem_Click(object sender, EventArgs e)
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
    }
}
