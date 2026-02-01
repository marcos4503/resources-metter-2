using CoroutinesDotNet;
using CoroutinesForWpf;
using LibreHardwareMonitor.Hardware;
using LibreHardwareMonitor.PawnIo;
using Resources_Metter_2.Scripts;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static Resources_Metter_2.Scripts.Preferences;

namespace Resources_Metter_2
{
    /*
     * This class holds and handles the central logic of Resources Metter 2 program.
    */

    public partial class MainWindow : Window
    {
        //Cache variables
        private bool isMoveModeEnabled = false;
        private KeyboardKeys_Watcher keyboardKeysWatcher = null;
        private KeyboardHotkey_Interceptor keyboardHotkeyCtrlUp = null;
        private KeyboardHotkey_Interceptor keyboardHotkeyCtrlRight = null;
        private KeyboardHotkey_Interceptor keyboardHotkeyCtrlLeft = null;
        private KeyboardHotkey_Interceptor keyboardHotkeyCtrlDown = null;
        private System.Windows.Forms.Timer windowsRestorerTimer = null;
        private SolidColorBrush progressBarDefaultColor = null;
        private SolidColorBrush progressBarAtMaxColor = null;
        private SolidColorBrush metricTextDefaultColor = null;
        private SolidColorBrush metricTextAtMaxColor = null;
        private bool fastMetricsInitialized = false;
        private bool normalMetricsInitialized = false;
        private bool slowMetricsInitialized = false;
        private bool netMetricsInitialized = false;
        private bool timeMetricsInitialized = false;
        private Process libreHardwareMonitorProcess = null;
        private WindowPercentView cpuUsageWindowPercentView = null;
        private WindowPercentView ramUsageWindowPercentView = null;
        private WindowPercentView gpuUsageWindowPercentView = null;
        private WindowPercentView vramUsageWindowPercentView = null;
        private WindowPercentView diskUsageWindowPercentView = null;

        //Private variables
        private WindowPositionDisplay windowPositionDisplay = null;
        private System.Windows.Forms.NotifyIcon notifyIcon = null;
        private System.Windows.Forms.ToolStripMenuItem notifyIcon_MoveModeOption = null;
        private System.Windows.Forms.ToolStripMenuItem notifyIcon_WindowsRestorerOption = null;
        private System.Windows.Forms.ToolStripMenuItem notifyIcon_LibreHMOption = null;
        private System.Windows.Forms.ToolStripMenuItem notifyIcon_AutoStartOption = null;
        private System.Windows.Forms.ToolStripMenuItem[] notifyIcon_cpuFanSlotButtons = null;
        private System.Windows.Forms.ToolStripMenuItem[] notifyIcon_cpuOptSlotButtons = null;
        private System.Windows.Forms.ToolStripMenuItem[] notifyIcon_cpuPumpSlotButtons = null;
        private Dictionary<string, System.Windows.Forms.ToolStripMenuItem> notifyIcon_pumpRpmWarnButtons = null;

        //Public variables
        public Preferences programPrefs = null;

        //Core methods

        public MainWindow()
        {
            //If the PawnIO is not installed, request the install and stop here
            if (PawnIo.IsInstalled == false)
            {
                //Warn abou the problem
                if (System.Windows.MessageBox.Show("The PawnIO driver is required to run the Resources Metter 2. Installation is quick, offline, and does not require a system restart. Do you want to install it now?", "PawnIO Required", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //Start the install proccess and wait for finish
                    Process installProcess = Process.Start(new ProcessStartInfo("Content/PawnIO_setup_2.0.1.exe", "-install"));
                    installProcess.WaitForExit();
                    System.Windows.MessageBox.Show("If the PawnIO installation was successful, please re-open Resources Meter 2.", "PawnIO Installation");
                }

                //Stop the execution of this instance
                System.Windows.Application.Current.Shutdown();

                //Cancel the execution
                return;
            }

            //Initialize the Window
            InitializeComponent();

            //Load the program preferences and get it
            programPrefs = new Preferences();

            //Inform the save informations and save it
            SaveInfo saveInfo = new SaveInfo();
            saveInfo.key = "saveVersion";
            saveInfo.value = "1.0.0";
            programPrefs.loadedData.saveInfo = new SaveInfo[] { saveInfo };
            programPrefs.Save();

            //Prepare the notify icon (Requires System.Drawing and System.Windows.Forms references on project)
            this.notifyIcon = new System.Windows.Forms.NotifyIcon();
            this.notifyIcon.Visible = true;
            this.notifyIcon.Text = "Resources Metter 2\nClick to pin to top!";
            this.notifyIcon.MouseClick += NotifyIcon_Click;
            this.notifyIcon.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            System.Windows.Forms.ToolStripMenuItem notifyIcon_UtilityFunctions = new ToolStripMenuItem("Functions For Utility", null, ((s, e) => { }));
            this.notifyIcon.ContextMenuStrip.Items.Add(notifyIcon_UtilityFunctions);
            this.notifyIcon_LibreHMOption = new System.Windows.Forms.ToolStripMenuItem("Libre Hardware Monitor", null, ((s, e) => { NotifyIcon_LibreHardwareMonitorOpen(); }));
            notifyIcon_UtilityFunctions.DropDownItems.Add(this.notifyIcon_LibreHMOption);
            this.notifyIcon_WindowsRestorerOption = new System.Windows.Forms.ToolStripMenuItem("Auto Restore Windows", null, ((s, e) => { NotifyIcon_WindowsRestorerToggle(); }));
            notifyIcon_UtilityFunctions.DropDownItems.Add(this.notifyIcon_WindowsRestorerOption);
            System.Windows.Forms.ToolStripMenuItem notifyIcon_PumpRpmWarn = new ToolStripMenuItem("Pump RPM Warn", null, ((s, e) => { }));
            this.notifyIcon.ContextMenuStrip.Items.Add(notifyIcon_PumpRpmWarn);
            notifyIcon_pumpRpmWarnButtons = new Dictionary<string, ToolStripMenuItem>();
            System.Windows.Forms.ToolStripMenuItem warnPumpRpmOff = new System.Windows.Forms.ToolStripMenuItem("Off", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(-1); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpmOff);
            notifyIcon_pumpRpmWarnButtons.Add("-1", warnPumpRpmOff);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm1500 = new System.Windows.Forms.ToolStripMenuItem("Below 1500rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(1500); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm1500);
            notifyIcon_pumpRpmWarnButtons.Add("1500", warnPumpRpm1500);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm1700 = new System.Windows.Forms.ToolStripMenuItem("Below 1700rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(1700); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm1700);
            notifyIcon_pumpRpmWarnButtons.Add("1700", warnPumpRpm1700);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm1900 = new System.Windows.Forms.ToolStripMenuItem("Below 1900rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(1900); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm1900);
            notifyIcon_pumpRpmWarnButtons.Add("1900", warnPumpRpm1900);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm2100 = new System.Windows.Forms.ToolStripMenuItem("Below 2100rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(2100); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm2100);
            notifyIcon_pumpRpmWarnButtons.Add("2100", warnPumpRpm2100);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm2300 = new System.Windows.Forms.ToolStripMenuItem("Below 2300rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(2300); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm2300);
            notifyIcon_pumpRpmWarnButtons.Add("2300", warnPumpRpm2300);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm2500 = new System.Windows.Forms.ToolStripMenuItem("Below 2500rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(2500); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm2500);
            notifyIcon_pumpRpmWarnButtons.Add("2500", warnPumpRpm2500);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm2700 = new System.Windows.Forms.ToolStripMenuItem("Below 2700rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(2700); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm2700);
            notifyIcon_pumpRpmWarnButtons.Add("2700", warnPumpRpm2700);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm2900 = new System.Windows.Forms.ToolStripMenuItem("Below 2900rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(2900); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm2900);
            notifyIcon_pumpRpmWarnButtons.Add("2900", warnPumpRpm2900);
            System.Windows.Forms.ToolStripMenuItem warnPumpRpm3000 = new System.Windows.Forms.ToolStripMenuItem("Below 3000rpm", null, ((s, e) => { NotifyIcon_ChangePumpRpmWarn(3000); }));
            notifyIcon_PumpRpmWarn.DropDownItems.Add(warnPumpRpm3000);
            notifyIcon_pumpRpmWarnButtons.Add("3000", warnPumpRpm3000);
            System.Windows.Forms.ToolStripMenuItem notifyIcon_SpecifyFanSlot = new ToolStripMenuItem("Specify Fan Slots", null, ((s, e) => { }));
            this.notifyIcon.ContextMenuStrip.Items.Add(notifyIcon_SpecifyFanSlot);
            System.Windows.Forms.ToolStripMenuItem notifyIcon_SpecifyFanSlotCpuFan = new ToolStripMenuItem("CPU Fan", null, ((s, e) => { }));
            notifyIcon_SpecifyFanSlot.DropDownItems.Add(notifyIcon_SpecifyFanSlotCpuFan);
            this.notifyIcon_cpuFanSlotButtons = new System.Windows.Forms.ToolStripMenuItem[10];
            for (int i = 0; i < 10; i++)
            {
                this.notifyIcon_cpuFanSlotButtons[i] = new System.Windows.Forms.ToolStripMenuItem(((i > 0) ? ("PWM #" + (i - 1)) : ("Off")), null, ((s, e) =>
                {
                    NotifyIcon_ChangeCpuFanPwmSlot((System.Windows.Forms.ToolStripMenuItem)s);
                }));
                notifyIcon_SpecifyFanSlotCpuFan.DropDownItems.Add(notifyIcon_cpuFanSlotButtons[i]);
            }
            System.Windows.Forms.ToolStripMenuItem notifyIcon_SpecifyFanSlotCpuOpt = new ToolStripMenuItem("CPU Opt", null, ((s, e) => { }));
            notifyIcon_SpecifyFanSlot.DropDownItems.Add(notifyIcon_SpecifyFanSlotCpuOpt);
            this.notifyIcon_cpuOptSlotButtons = new System.Windows.Forms.ToolStripMenuItem[10];
            for (int i = 0; i < 10; i++)
            {
                this.notifyIcon_cpuOptSlotButtons[i] = new System.Windows.Forms.ToolStripMenuItem(((i > 0) ? ("PWM #" + (i - 1)) : ("Off")), null, ((s, e) =>
                {
                    NotifyIcon_ChangeCpuOptPwmSlot((System.Windows.Forms.ToolStripMenuItem)s);
                }));
                notifyIcon_SpecifyFanSlotCpuOpt.DropDownItems.Add(notifyIcon_cpuOptSlotButtons[i]);
            }
            System.Windows.Forms.ToolStripMenuItem notifyIcon_SpecifyFanSlotCpuPump = new ToolStripMenuItem("CPU Pump", null, ((s, e) => { }));
            notifyIcon_SpecifyFanSlot.DropDownItems.Add(notifyIcon_SpecifyFanSlotCpuPump);
            this.notifyIcon_cpuPumpSlotButtons = new System.Windows.Forms.ToolStripMenuItem[10];
            for (int i = 0; i < 10; i++)
            {
                this.notifyIcon_cpuPumpSlotButtons[i] = new System.Windows.Forms.ToolStripMenuItem(((i > 0) ? ("PWM #" + (i - 1)) : ("Off")), null, ((s, e) =>
                {
                    NotifyIcon_ChangeCpuPumpPwmSlot((System.Windows.Forms.ToolStripMenuItem)s);
                }));
                notifyIcon_SpecifyFanSlotCpuPump.DropDownItems.Add(notifyIcon_cpuPumpSlotButtons[i]);
            }
            this.notifyIcon_AutoStartOption = new System.Windows.Forms.ToolStripMenuItem("Start With System", null, ((s, e) => { NotifyIcon_StartWithSystemToggle(); }));
            this.notifyIcon.ContextMenuStrip.Items.Add(this.notifyIcon_AutoStartOption);
            this.notifyIcon_MoveModeOption = new System.Windows.Forms.ToolStripMenuItem("Move Mode", null, ((s, e) => { NotifyIcon_MoveModeToggle(); }));
            this.notifyIcon.ContextMenuStrip.Items.Add(this.notifyIcon_MoveModeOption);
            this.notifyIcon.ContextMenuStrip.Items.Add("Quit", null, ((s, e) => { QuitApp(); }));
            this.notifyIcon.Icon = new System.Drawing.Icon(@"Content/tray-icon-normal.ico");

            //Restore the position of program
            this.Loaded += (s, e) =>
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Left = programPrefs.loadedData.windowPosX;
                this.Top = programPrefs.loadedData.windowPosY;
            };
            //Restore the auto restore windows
            if (programPrefs.loadedData.autoRestoreWindows == true)
                EnableTheAutoRestoreWindowsTimer();
            notifyIcon_WindowsRestorerOption.Checked = programPrefs.loadedData.autoRestoreWindows;
            //Restore the current selection of PWM slot for CPU Fan
            notifyIcon_cpuFanSlotButtons[(programPrefs.loadedData.cpuFanPwmSlot + 1)].Checked = true;
            //Restore the current selection of PWM slot for CPU Opt
            notifyIcon_cpuOptSlotButtons[(programPrefs.loadedData.cpuOptPwmSlot + 1)].Checked = true;
            //Restore the current selection of PWM slot for CPU Pump
            notifyIcon_cpuPumpSlotButtons[(programPrefs.loadedData.cpuPumpPwmSlot + 1)].Checked = true;
            //Restore the current selection of Pump RPM Warn
            notifyIcon_pumpRpmWarnButtons[programPrefs.loadedData.cpuPumpRpmWarn.ToString()].Checked = true;
            //Restore the current selection of auto start with system
            notifyIcon_AutoStartOption.Checked = programPrefs.loadedData.autoStartWithSystem;

            //Prepare click to view interactions
            cpuUsageProgressBar.MouseLeftButtonUp += ((s, e) => { ClickToViewCpuUsage(); });
            ramUsageProgressBar.MouseLeftButtonUp += ((s, e) => { ClickToViewRamUsage(); });
            gpuUsageProgressBar.MouseLeftButtonUp += ((s, e) => { ClickToViewGpuUsage(); });
            vramUsageProgressBar.MouseLeftButtonUp += ((s, e) => { ClickToViewVramUsage(); });
            diskUsageProgressBar.MouseLeftButtonUp += ((s, e) => { ClickToViewDiskUsage(); });

            //Initialize the metrics renderization
            InitializeMetricsRenderization();

            //Start the routine to wait initialization of all metrics
            IDisposable metricsInitializationRoutine = Coroutine.Start(WaitAllInitializationsRoutine());
        }

        private void InitializeMetricsRenderization()
        {
            progressBarDefaultColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 139, 165));
            progressBarAtMaxColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 208, 0, 0));
            metricTextDefaultColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 0, 0));
            metricTextAtMaxColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 206, 0, 0));

            //Create a slow thread to render the metrics
            Thread slowMetricsThread = new Thread(() =>
            {
                //Initialize the RAM usage reader
                Microsoft.VisualBasic.Devices.ComputerInfo computerInfo = new Microsoft.VisualBasic.Devices.ComputerInfo();
                float totalRamMb = computerInfo.TotalPhysicalMemory / (1024 * 1024);
                PerformanceCounter ramUsageReader = new PerformanceCounter("Memory", "Available MBytes");
                float lastRamNextValue = ramUsageReader.NextValue();
                //Initialize the LibreHardwareMonitorLib computer for slow reading
                Computer computerReaderSlow = new Computer { IsMotherboardEnabled = true };
                computerReaderSlow.Open();
                computerReaderSlow.Accept(new UpdateVisitor());

                //Prepare cache variables
                int cpuFanRpm = 0;
                int cpuOptRpm = 0;
                int cpuPumpRpm = 0;
                float ramUsageNow = 0;

                //Get UI cache data
                System.Windows.Threading.Dispatcher currentApplicationDispatcher = System.Windows.Application.Current.Dispatcher;
                Action actionToRunOnUI = () =>
                {
                    //If the CPU Pump display is disabled
                    if (programPrefs.loadedData.cpuPumpPwmSlot == -1)
                    {
                        if (pumpImg.Visibility != Visibility.Collapsed)
                        {
                            pumpTxt.Visibility = Visibility.Collapsed;
                            pumpImg.Visibility = Visibility.Collapsed;
                        }
                    }
                    //If the CPU Pump display is enabled
                    if (programPrefs.loadedData.cpuPumpPwmSlot != -1)
                    {
                        if (pumpImg.Visibility != Visibility.Visible)
                        {
                            pumpTxt.Visibility = Visibility.Visible;
                            pumpImg.Visibility = Visibility.Visible;
                        }
                        //Render the CPU Pump RPM
                        if (cpuPumpRpm == 0)
                            pumpTxt.Text = "Disabled";
                        if (cpuPumpRpm > 0)
                            pumpTxt.Text = (cpuPumpRpm + "rpm");
                        if (programPrefs.loadedData.cpuPumpRpmWarn != -1)
                        {
                            if (cpuPumpRpm < programPrefs.loadedData.cpuPumpRpmWarn)
                                pumpTxt.Foreground = metricTextAtMaxColor;
                            if (cpuPumpRpm >= programPrefs.loadedData.cpuPumpRpmWarn)
                                pumpTxt.Foreground = metricTextDefaultColor;
                        }
                        if (programPrefs.loadedData.cpuPumpRpmWarn == -1)
                            pumpTxt.Foreground = metricTextDefaultColor;
                    }

                    //Calculate the resume CPU RPM to show
                    int cpuRpmToDisplay = -1;
                    if (programPrefs.loadedData.cpuFanPwmSlot == -1 && programPrefs.loadedData.cpuOptPwmSlot == -1)
                        cpuRpmToDisplay = -1;
                    if (programPrefs.loadedData.cpuFanPwmSlot != -1 && programPrefs.loadedData.cpuOptPwmSlot == -1)
                        cpuRpmToDisplay = cpuFanRpm;
                    if (programPrefs.loadedData.cpuFanPwmSlot == -1 && programPrefs.loadedData.cpuOptPwmSlot != -1)
                        cpuRpmToDisplay = cpuOptRpm;
                    if (programPrefs.loadedData.cpuFanPwmSlot != -1 && programPrefs.loadedData.cpuOptPwmSlot != -1)
                        cpuRpmToDisplay = (int)(((float)cpuFanRpm + (float)cpuOptRpm) / 2.0f);
                    if (cpuRpmToDisplay == -1)
                        cpuFanTxt.Text = "Off";
                    if (cpuRpmToDisplay != -1)
                    {
                        if (cpuRpmToDisplay == 0)
                            cpuFanTxt.Text = "Disabled";
                        if (cpuRpmToDisplay > 0)
                            cpuFanTxt.Text = (cpuRpmToDisplay + "rpm");
                    }

                    //Display the RAM usage
                    ramUsageProgressBar.Value = ramUsageNow;
                    if (ramUsageNow >= 90)
                        ramUsageProgressBar.Foreground = progressBarAtMaxColor;
                    if (ramUsageNow < 90)
                        ramUsageProgressBar.Foreground = progressBarDefaultColor;
                    //If the Percent Window View is present, update this too...
                    if (ramUsageWindowPercentView != null)
                    {
                        ramUsageWindowPercentView.value.Text = (ramUsageNow.ToString("F0") + "%");
                        if (ramUsageNow >= 90)
                            ramUsageWindowPercentView.value.Foreground = metricTextAtMaxColor;
                        if (ramUsageNow < 90)
                            ramUsageWindowPercentView.value.Foreground = metricTextDefaultColor;
                    }
                };

                //Run a loop to update the metrics
                while (true)
                {
                    //Get RAM usage
                    float ramInUse = totalRamMb - ramUsageReader.NextValue();
                    float ramUsageNextValue = ramInUse / totalRamMb * 100.0f;
                    ramUsageNow = ramUsageNextValue;

                    //Read the computer using LibreHardwareMonitorLib
                    foreach (IHardware hardware in computerReaderSlow.Hardware)
                    {
                        //Read the Motherboard
                        if (hardware.HardwareType == HardwareType.Motherboard)
                        {
                            //Update the hardware info
                            hardware.Update();

                            //Read each sub-hardware of this Motherboard
                            foreach (IHardware subHardware in hardware.SubHardware)
                            {
                                //Update the sub-hardware info
                                subHardware.Update();

                                //Prepare the cpu pump variable
                                float? cpuPumpRpmTmp = -2;
                                //Try to get the CPU Pump RPM
                                foreach (ISensor sensor in subHardware.Sensors)
                                    if (sensor.SensorType == SensorType.Fan && sensor.Name.Contains(programPrefs.loadedData.cpuPumpPwmSlot.ToString()))
                                        if (sensor.Value != null)
                                            cpuPumpRpmTmp = sensor.Value;
                                //Store the value
                                if (cpuPumpRpmTmp != -2)
                                {
                                    if (programPrefs.loadedData.cpuPumpPwmSlot == -1)
                                        cpuPumpRpm = -1;
                                    if (programPrefs.loadedData.cpuPumpPwmSlot != -1)
                                        cpuPumpRpm = (int)(float)cpuPumpRpmTmp;
                                }

                                //Prepare the cpu fan variable
                                float? cpuFanRpmTmp = -2;
                                //Try to get the CPU Fan RPM
                                foreach (ISensor sensor in subHardware.Sensors)
                                    if (sensor.SensorType == SensorType.Fan && sensor.Name.Contains(programPrefs.loadedData.cpuFanPwmSlot.ToString()))
                                        if (sensor.Value != null)
                                            cpuFanRpmTmp = sensor.Value;
                                //Store the value
                                if (cpuFanRpmTmp != -2)
                                {
                                    if (programPrefs.loadedData.cpuFanPwmSlot == -1)
                                        cpuFanRpm = -1;
                                    if (programPrefs.loadedData.cpuFanPwmSlot != -1)
                                        cpuFanRpm = (int)(float)cpuFanRpmTmp;
                                }

                                //Prepare the cpu opt variable
                                float? cpuOptRpmTmp = -2;
                                //Try to get the CPU Opt RPM
                                foreach (ISensor sensor in subHardware.Sensors)
                                    if (sensor.SensorType == SensorType.Fan && sensor.Name.Contains(programPrefs.loadedData.cpuOptPwmSlot.ToString()))
                                        if (sensor.Value != null)
                                            cpuOptRpmTmp = sensor.Value;
                                //Store the value
                                if (cpuOptRpmTmp != -2)
                                {
                                    if (programPrefs.loadedData.cpuOptPwmSlot == -1)
                                        cpuOptRpm = -1;
                                    if (programPrefs.loadedData.cpuOptPwmSlot != -1)
                                        cpuOptRpm = (int)(float)cpuOptRpmTmp;
                                }
                            }
                        }
                    }

                    //Update the UI
                    currentApplicationDispatcher.Invoke(actionToRunOnUI);

                    //Wait the interval
                    Thread.Sleep(1000);

                    //Inform that was initialized
                    slowMetricsInitialized = true;
                }
            });
            slowMetricsThread.Start();

            //Create a normal thread to render the metrics
            Thread normalMetricsThread = new Thread(() =>
            {
                //Initialize the Disk usage reader
                PerformanceCounter diskUsageReader = new PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total");
                float lastDiskNextValue = diskUsageReader.NextValue();

                //Prepare cache variables
                float diskUsageNow = 0;

                //Get UI cache data
                System.Windows.Threading.Dispatcher currentApplicationDispatcher = System.Windows.Application.Current.Dispatcher;
                Action actionToRunOnUI = () =>
                {
                    //Render the Disk Usage
                    diskUsageProgressBar.Value = diskUsageNow;
                    if (diskUsageNow >= 90)
                        diskUsageProgressBar.Foreground = progressBarAtMaxColor;
                    if (diskUsageNow < 90)
                        diskUsageProgressBar.Foreground = progressBarDefaultColor;
                    //If the Percent Window View is present, update this too...
                    if (diskUsageWindowPercentView != null)
                    {
                        diskUsageWindowPercentView.value.Text = (diskUsageNow.ToString("F0") + "%");
                        if (diskUsageNow >= 90)
                            diskUsageWindowPercentView.value.Foreground = metricTextAtMaxColor;
                        if (diskUsageNow < 90)
                            diskUsageWindowPercentView.value.Foreground = metricTextDefaultColor;
                    }
                };

                //Run a loop to update the metrics
                while (true)
                {
                    diskUsageNow = diskUsageReader.NextValue();

                    //Update the UI
                    currentApplicationDispatcher.Invoke(actionToRunOnUI);

                    //Wait the interval
                    Thread.Sleep(500);

                    //Inform that was initialized
                    normalMetricsInitialized = true;
                }
            });
            normalMetricsThread.Start();

            //Create a fast thread to render the metrics
            Thread fastMetricsThread = new Thread(() =>
            {
                //Initialize the CPU usage reader
                PerformanceCounter cpuUsageReader = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                float lastCpuNextValue = cpuUsageReader.NextValue();
                //Initialize the LibreHardwareMonitorLib computer for fast reading
                Computer computerReaderFast = new Computer { IsCpuEnabled = true, IsGpuEnabled = true };
                computerReaderFast.Open();
                computerReaderFast.Accept(new UpdateVisitor());

                //Prepare cache variables
                float cpuUsage1 = 0;
                float cpuUsage0 = 0;
                float cpuUsageNow = 0;
                float cpuTemp3 = 0;
                float cpuTemp2 = 0;
                float cpuTemp1 = 0;
                float cpuTemp0 = 0;
                float cpuTempNow = 0;
                float cpuTdp8 = 0;
                float cpuTdp7 = 0;
                float cpuTdp6 = 0;
                float cpuTdp5 = 0;
                float cpuTdp4 = 0;
                float cpuTdp3 = 0;
                float cpuTdp2 = 0;
                float cpuTdp1 = 0;
                float cpuTdp0 = 0;
                float cpuTdpNow = 0;
                float gpuUsageNow = 0;
                float gpuTempNow = 0;
                float gpuTdp3 = 0;
                float gpuTdp2 = 0;
                float gpuTdp1 = 0;
                float gpuTdp0 = 0;
                float gpuTdpNow = 0;
                int gpuFan1Rpm = 0;
                int gpuFan2Rpm = 0;
                int gpuFan3Rpm = 0;
                float vramUsageNow = 0;
                float vramTempNow = 0;

                //Get UI cache data
                System.Windows.Threading.Dispatcher currentApplicationDispatcher = System.Windows.Application.Current.Dispatcher;
                Action actionToRunOnUI = () =>
                {
                    //Render the CPU Usage
                    float cpuUseAvg = ((cpuUsageNow + cpuUsage0) / 2.0f);
                    cpuUsageProgressBar.Value = cpuUseAvg;
                    if (cpuUseAvg >= 90)
                        cpuUsageProgressBar.Foreground = progressBarAtMaxColor;
                    if (cpuUseAvg < 90)
                        cpuUsageProgressBar.Foreground = progressBarDefaultColor;
                    //If the Percent Window View is present, update this too...
                    if (cpuUsageWindowPercentView != null)
                    {
                        cpuUsageWindowPercentView.value.Text = (cpuUseAvg.ToString("F0") + "%");
                        if (cpuUseAvg >= 90)
                            cpuUsageWindowPercentView.value.Foreground = metricTextAtMaxColor;
                        if (cpuUseAvg < 90)
                            cpuUsageWindowPercentView.value.Foreground = metricTextDefaultColor;
                    }

                    //Render the CPU Temperature
                    int cpuTempAvg = (int)((cpuTempNow + cpuTemp0 + cpuTemp1 + cpuTemp2 + cpuTemp3) / 5.0f);
                    cpuTempTxt.Text = (cpuTempAvg + "°");
                    if (cpuTempAvg >= 80)
                        cpuTempTxt.Foreground = metricTextAtMaxColor;
                    if (cpuTempAvg < 80)
                        cpuTempTxt.Foreground = metricTextDefaultColor;

                    //Render the CPU TDP
                    int cpuTdpAvg = (int)((cpuTdpNow + cpuTdp0 + cpuTdp1 + cpuTdp2 + cpuTdp3 + cpuTdp4 + cpuTdp5 + cpuTdp6 + cpuTdp7 + cpuTdp8) / 10.0f);
                    cpuTdpTxt.Text = (cpuTdpAvg + "w");

                    //Render the GPU Usage
                    gpuUsageProgressBar.Value = gpuUsageNow;
                    if (gpuUsageNow >= 90)
                        gpuUsageProgressBar.Foreground = progressBarAtMaxColor;
                    if (gpuUsageNow < 90)
                        gpuUsageProgressBar.Foreground = progressBarDefaultColor;
                    //If the Percent Window View is present, update this too...
                    if (gpuUsageWindowPercentView != null)
                    {
                        gpuUsageWindowPercentView.value.Text = (gpuUsageNow.ToString("F0") + "%");
                        if (gpuUsageNow >= 90)
                            gpuUsageWindowPercentView.value.Foreground = metricTextAtMaxColor;
                        if (gpuUsageNow < 90)
                            gpuUsageWindowPercentView.value.Foreground = metricTextDefaultColor;
                    }

                    //Render the GPU Temperature
                    gpuTempTxt.Text = (gpuTempNow.ToString("F0") + "°");
                    if (gpuTempNow >= 80)
                        gpuTempTxt.Foreground = metricTextAtMaxColor;
                    if (gpuTempNow < 80)
                        gpuTempTxt.Foreground = metricTextDefaultColor;

                    //Render the GPU TDP
                    int gpuTdpAvg = (int)((gpuTdpNow + gpuTdp0 + gpuTdp1 + gpuTdp2 + gpuTdp3) / 5.0f);
                    gpuTdpTxt.Text = (gpuTdpAvg + "w");

                    //Calculate the resume GPU RPM to show
                    int gpuRpmToDisplay = -1;
                    if (gpuFan1Rpm == -1 && gpuFan2Rpm == -1 && gpuFan3Rpm == -1)
                        gpuRpmToDisplay = -1;
                    if (gpuFan1Rpm != -1 && gpuFan2Rpm == -1 && gpuFan3Rpm == -1)
                        gpuRpmToDisplay = gpuFan1Rpm;
                    if (gpuFan1Rpm != -1 && gpuFan2Rpm != -1 && gpuFan3Rpm == -1)
                        gpuRpmToDisplay = (int)(((float)gpuFan1Rpm + (float)gpuFan2Rpm) / 2.0f);
                    if (gpuFan1Rpm != -1 && gpuFan2Rpm != -1 && gpuFan3Rpm != -1)
                        gpuRpmToDisplay = (int)(((float)gpuFan1Rpm + (float)gpuFan2Rpm + (float)gpuFan3Rpm) / 3.0f);
                    if (gpuRpmToDisplay == -1)
                        gpuFanTxt.Text = "Off";
                    if (gpuRpmToDisplay != -1)
                    {
                        if (gpuRpmToDisplay == 0)
                            gpuFanTxt.Text = "Disabled";
                        if (gpuRpmToDisplay > 0)
                            gpuFanTxt.Text = (gpuRpmToDisplay + "rpm");
                    }

                    //Render the VRAM usage
                    vramUsageProgressBar.Value = vramUsageNow;
                    if (vramUsageNow >= 90)
                        vramUsageProgressBar.Foreground = progressBarAtMaxColor;
                    if (vramUsageNow < 90)
                        vramUsageProgressBar.Foreground = progressBarDefaultColor;
                    //If the Percent Window View is present, update this too...
                    if (vramUsageWindowPercentView != null)
                    {
                        vramUsageWindowPercentView.value.Text = (vramUsageNow.ToString("F0") + "%");
                        if (vramUsageNow >= 90)
                            vramUsageWindowPercentView.value.Foreground = metricTextAtMaxColor;
                        if (vramUsageNow < 90)
                            vramUsageWindowPercentView.value.Foreground = metricTextDefaultColor;
                    }

                    //Render the VRAM Temperature
                    vramTempTxt.Text = (vramTempNow + "°");
                    if (vramTempNow >= 90)
                        vramTempTxt.Foreground = metricTextAtMaxColor;
                    if (vramTempNow < 90)
                        vramTempTxt.Foreground = metricTextDefaultColor;
                };

                //Run a loop to update the metrics
                while (true)
                {
                    //Get the CPU usage
                    cpuUsage1 = cpuUsage0;
                    cpuUsage0 = cpuUsageNow;
                    cpuUsageNow = cpuUsageReader.NextValue();

                    //Read the computer using LibreHardwareMonitorLib
                    foreach (IHardware hardware in computerReaderFast.Hardware)
                    {
                        //Read the CPU
                        if (hardware.HardwareType == HardwareType.Cpu)
                        {
                            //Update the hardware info
                            hardware.Update();

                            //Prepare the temp variable
                            float? cpuTmp = 0;
                            //Try to get the CPU temperature
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Temperature && sensor.Name == "CPU Package")
                                    if (sensor.Value != null)
                                        cpuTmp = sensor.Value;
                            //If not found a value in "CPU Package" sensor, try to found in "Core (Tctl/Tdie)". This is useful for AMD CPUs
                            if (cpuTmp == 0)
                                foreach (ISensor sensor in hardware.Sensors)
                                    if (sensor.SensorType == SensorType.Temperature && sensor.Name == "Core (Tctl/Tdie)")
                                        if (sensor.Value != null)
                                            cpuTmp = sensor.Value;
                            //Store the value
                            cpuTemp3 = cpuTemp2;
                            cpuTemp2 = cpuTemp1;
                            cpuTemp1 = cpuTemp0;
                            cpuTemp0 = cpuTempNow;
                            cpuTempNow = (float)cpuTmp;

                            //Prepare the TDP variable
                            float? cpuTdp = 0;
                            //Try to get the CPU TDP
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Power && sensor.Name == "CPU Package")
                                    if (sensor.Value != null)
                                        cpuTdp = sensor.Value;
                            //If not found a value in "CPU Package" sensor, try to found in "Core (Tctl/Tdie)". This is useful for AMD CPUs
                            if (cpuTdp == 0)
                                foreach (ISensor sensor in hardware.Sensors)
                                    if (sensor.SensorType == SensorType.Power && sensor.Name == "Core (Tctl/Tdie)")
                                        if (sensor.Value != null)
                                            cpuTdp = sensor.Value;
                            //Store the value
                            cpuTdp8 = cpuTdp7;
                            cpuTdp7 = cpuTdp6;
                            cpuTdp6 = cpuTdp5;
                            cpuTdp5 = cpuTdp4;
                            cpuTdp4 = cpuTdp3;
                            cpuTdp3 = cpuTdp2;
                            cpuTdp2 = cpuTdp1;
                            cpuTdp1 = cpuTdp0;
                            cpuTdp0 = cpuTdpNow;
                            cpuTdpNow = (float)cpuTdp;
                        }

                        //Read the GPU
                        if (hardware.HardwareType == HardwareType.GpuNvidia || hardware.HardwareType == HardwareType.GpuAmd)
                        {
                            //Update the hardware info
                            hardware.Update();

                            //Prepare the usage variable
                            float? gpuUsage = 0;
                            //Try to get the GPU usage
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Load && sensor.Name == "GPU Core")
                                    if (sensor.Value != null)
                                        gpuUsage = sensor.Value;
                            //Store the value
                            gpuUsageNow = (float)gpuUsage;

                            //Prepare the temp variable
                            float? gpuTemp = 0;
                            //Try to get the GPU temperature
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Temperature && sensor.Name == "GPU Core")
                                    if (sensor.Value != null)
                                        gpuTemp = sensor.Value;
                            //Store the value
                            gpuTempNow = (float)gpuTemp;

                            //Prepare the tdp variable
                            float? gpuTdp = 0;
                            //Try to get the GPU tdp
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Power && sensor.Name == "GPU Package")
                                    if (sensor.Value != null)
                                        gpuTdp = sensor.Value;
                            //Store the value
                            gpuTdp3 = gpuTdp2;
                            gpuTdp2 = gpuTdp1;
                            gpuTdp1 = gpuTdp0;
                            gpuTdp0 = gpuTdpNow;
                            gpuTdpNow = (float)gpuTdp;

                            //Prepare the gpu fan 1 rpm
                            float? gpuFan1RpmTmp = -1;
                            //Try to get the GPU fan 1 rpm
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Fan && sensor.Name == "GPU Fan 1")
                                    if (sensor.Value != null)
                                        gpuFan1RpmTmp = sensor.Value;
                            if (gpuFan1RpmTmp == -1)
                                foreach (ISensor sensor in hardware.Sensors)
                                    if (sensor.SensorType == SensorType.Fan && sensor.Name == "GPU Fan")
                                        if (sensor.Value != null)
                                            gpuFan1RpmTmp = sensor.Value;
                            //Store the value
                            gpuFan1Rpm = (int)(float)gpuFan1RpmTmp;

                            //Prepare the gpu fan 2 rpm
                            float? gpuFan2RpmTmp = -1;
                            //Try to get the GPU fan 2 rpm
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Fan && sensor.Name == "GPU Fan 2")
                                    if (sensor.Value != null)
                                        gpuFan2RpmTmp = sensor.Value;
                            //Store the value
                            gpuFan2Rpm = (int)(float)gpuFan2RpmTmp;

                            //Prepare the gpu fan 3 rpm
                            float? gpuFan3RpmTmp = -1;
                            //Try to get the GPU fan 3 rpm
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Fan && sensor.Name == "GPU Fan 3")
                                    if (sensor.Value != null)
                                        gpuFan3RpmTmp = sensor.Value;
                            //Store the value
                            gpuFan3Rpm = (int)(float)gpuFan3RpmTmp;

                            //Prepare the vram usage variable
                            float? vramUsage = 0;
                            //Try to get the GPU VRAM usage
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Load && sensor.Name == "GPU Memory")
                                    if (sensor.Value != null)
                                        vramUsage = sensor.Value;
                            //Store the value
                            vramUsageNow = (float)vramUsage;

                            //Prepare the vram temp variable
                            float? vramTemp = 0;
                            //Try to get the GPU VRAM temp
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Temperature && sensor.Name == "GPU Memory Junction")
                                    if (sensor.Value != null)
                                        vramTemp = sensor.Value;
                            //Store the value
                            vramTempNow = (float)vramTemp;
                        }
                    }

                    //Update the UI
                    currentApplicationDispatcher.Invoke(actionToRunOnUI);

                    //Wait the interval
                    Thread.Sleep(200);

                    //Inform that was initialized
                    fastMetricsInitialized = true;
                }
            });
            fastMetricsThread.Start();

            //Create a network thread to render the network usage
            Thread netMetricsThread = new Thread(() =>
            {
                //Initialize the LibreHardwareMonitorLib computer for network reading
                Computer computerReaderNet = new Computer { IsNetworkEnabled = true };
                computerReaderNet.Open();
                computerReaderNet.Accept(new UpdateVisitor());

                //Prepare cache variables
                float netDownload = 0;
                float netUpload = 0;

                //Get UI cache data
                System.Windows.Threading.Dispatcher currentApplicationDispatcher = System.Windows.Application.Current.Dispatcher;
                Action actionToRunOnUI = () =>
                {
                    //Prepare the minimum speed in KibiBytes to show the arrow accent
                    int minSpeedToShowArrowAccentInKb = 1000;

                    //Render the current download speed in MebiBytes and KiBibytes
                    float downloadMiB = ((netDownload / 1024.0f) / 1024.0f);
                    float downloadKiB = ((netDownload / 1024.0f));
                    string downloadStr = "";
                    if (downloadKiB <= 999)
                    {
                        downloadStr = (downloadKiB.ToString("F0") + " KB/s");
                    }
                    if (downloadKiB > 999)
                    {
                        //if (downloadMiB < 10)
                        //    downloadStr = ("0" + downloadMiB.ToString("F1") + " MB/s");
                        //if (downloadMiB >= 10)
                        downloadStr = (downloadMiB.ToString("F1") + " MB/s");
                    }
                    if (downloadKiB >= minSpeedToShowArrowAccentInKb)
                        if (downloadAccent.Visibility != Visibility.Visible)
                            downloadAccent.Visibility = Visibility.Visible;
                    if (downloadKiB < minSpeedToShowArrowAccentInKb)
                        if (downloadAccent.Visibility != Visibility.Collapsed)
                            downloadAccent.Visibility = Visibility.Collapsed;
                    downloadTxt.Text = downloadStr;

                    //Render the current upload speed in MebiBytes and KibiBytes
                    float uploadMiB = ((netUpload / 1024.0f) / 1024.0f);
                    float uploadKiB = ((netUpload / 1024.0f));
                    string uploadStr = "";
                    if (uploadKiB <= 999)
                    {
                        uploadStr = (uploadKiB.ToString("F0") + " KB/s");
                    }
                    if (uploadKiB > 999)
                    {
                        //if (uploadMiB < 10)
                        //    uploadStr = ("0" + uploadMiB.ToString("F1") + " MB/s");
                        //if (uploadMiB >= 10)
                        uploadStr = (uploadMiB.ToString("F1") + " MB/s");
                    }
                    if (uploadKiB >= minSpeedToShowArrowAccentInKb)
                        if (uploadAccent.Visibility != Visibility.Visible)
                            uploadAccent.Visibility = Visibility.Visible;
                    if (uploadKiB < minSpeedToShowArrowAccentInKb)
                        if (uploadAccent.Visibility != Visibility.Collapsed)
                            uploadAccent.Visibility = Visibility.Collapsed;
                    uploadTxt.Text = uploadStr;
                };

                //Run a loop to update the metrics
                while (true)
                {
                    //Set the download and upload to zero, until find Network hardware
                    netDownload = 0;
                    netUpload = 0;

                    //Read the computer using LibreHardwareMonitorLib
                    foreach (IHardware hardware in computerReaderNet.Hardware)
                    {
                        //Read the Network Adapter
                        if (hardware.HardwareType == HardwareType.Network)
                        {
                            //Update the hardware info
                            hardware.Update();

                            //Prepare the download variable
                            float? downloadTemp = 0;
                            //Try to get the download speed
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Throughput && sensor.Name == "Download Speed")
                                    if (sensor.Value != null)
                                        downloadTemp = sensor.Value;
                            //Store the value
                            netDownload = (float)downloadTemp;

                            //Prepare the upload variable
                            float? uploadTemp = 0;
                            //Try to get the upload speed
                            foreach (ISensor sensor in hardware.Sensors)
                                if (sensor.SensorType == SensorType.Throughput && sensor.Name == "Upload Speed")
                                    if (sensor.Value != null)
                                        uploadTemp = sensor.Value;
                            //Store the value
                            netUpload = (float)uploadTemp;
                        }
                    }

                    //Update the UI
                    currentApplicationDispatcher.Invoke(actionToRunOnUI);

                    //Wait the interval
                    Thread.Sleep(300);

                    //Inform that was initialized
                    netMetricsInitialized = true;
                }
            });
            netMetricsThread.Start();

            //Create the time thread to render the time
            Thread timeMetricsThread = new Thread(() =>
            {
                //Prepare cache variables
                string timeTxt = "--:--";

                //Get UI cache data
                System.Windows.Threading.Dispatcher currentApplicationDispatcher = System.Windows.Application.Current.Dispatcher;
                Action actionToRunOnUI = () =>
                {
                    //Render the time
                    sysTimeTxt.Text = timeTxt;
                };

                //Run a loop to update the time
                while (true)
                {
                    //Get current system time
                    DateTime dateTime = DateTime.Now;
                    int hour = dateTime.Hour;
                    int minute = dateTime.Minute;
                    timeTxt = ((hour > 9) ? hour.ToString() : "0" + hour) + ":" + ((minute > 9) ? minute.ToString() : "0" + minute);

                    //Update the UI
                    currentApplicationDispatcher.Invoke(actionToRunOnUI);

                    //Inform that was initialized
                    timeMetricsInitialized = true;

                    //Wait the interval
                    Thread.Sleep(30000);
                }
            });
            timeMetricsThread.Start();
        }

        private IEnumerator WaitAllInitializationsRoutine()
        {
            //Show the loading indicator
            loadIndicator.Visibility = Visibility.Visible;
            bg.Visibility = Visibility.Hidden;

            //Prepare the wait interval
            WaitForSeconds loopInterval = new WaitForSeconds(0.1f);

            //Wait until all is initialized
            while (true)
            {
                //Wait the interval of loop
                yield return loopInterval;

                //If all is initialized, break this loop
                if (fastMetricsInitialized == true && normalMetricsInitialized == true && slowMetricsInitialized == true && netMetricsInitialized == true && timeMetricsInitialized == true)
                    break;
            }

            //Wait time
            yield return new WaitForSeconds(1.0f);

            //Disable the loading indicator
            loadIndicator.Visibility = Visibility.Collapsed;
            bg.Visibility = Visibility.Visible;
        }

        //Auxiliar methods

        private void NotifyIcon_Click(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            //If not is left click, skip
            if (e.Button != MouseButtons.Left)
                return;

            //Toggle the topmost on click in Notify Icon
            if (this.Topmost == true)
            {
                //Change to default icon
                this.notifyIcon.Icon = new System.Drawing.Icon(@"Content/tray-icon-normal.ico");

                //Change to default mode for the main window
                this.Topmost = false;

                //Change to default mode for all percent view windows that exists
                if (cpuUsageWindowPercentView != null)
                    cpuUsageWindowPercentView.Topmost = false;
                if (ramUsageWindowPercentView != null)
                    ramUsageWindowPercentView.Topmost = false;
                if (gpuUsageWindowPercentView != null)
                    gpuUsageWindowPercentView.Topmost = false;
                if (vramUsageWindowPercentView != null)
                    vramUsageWindowPercentView.Topmost = false;
                if (diskUsageWindowPercentView != null)
                    diskUsageWindowPercentView.Topmost = false;

                //Cancel here
                return;
            }
            if (this.Topmost == false)
            {
                //Change to pinned icon
                this.notifyIcon.Icon = new System.Drawing.Icon(@"Content/tray-icon-pinned.ico");

                //Change to pinned mode for the main window
                this.Topmost = true;

                //Change to pinned mode for all percent view windows that exists
                if (cpuUsageWindowPercentView != null)
                    cpuUsageWindowPercentView.Topmost = true;
                if (ramUsageWindowPercentView != null)
                    ramUsageWindowPercentView.Topmost = true;
                if (gpuUsageWindowPercentView != null)
                    gpuUsageWindowPercentView.Topmost = true;
                if (vramUsageWindowPercentView != null)
                    vramUsageWindowPercentView.Topmost = true;
                if (diskUsageWindowPercentView != null)
                    diskUsageWindowPercentView.Topmost = true;

                //Cancel here
                return;
            }
        }

        private void NotifyIcon_MoveModeToggle()
        {
            //If the move mode is enabled...
            if (isMoveModeEnabled == true)
            {
                //Change the UI to default mode
                bg.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 95, 127));

                //Close the position display window
                if (windowPositionDisplay != null)
                {
                    windowPositionDisplay.Close();
                    windowPositionDisplay = null;
                }

                //Set the move mode option checked
                notifyIcon_MoveModeOption.Checked = false;

                //Disable the key listeners
                if (keyboardKeysWatcher != null)
                {
                    keyboardKeysWatcher.Dispose();
                    keyboardKeysWatcher = null;
                }
                if (keyboardHotkeyCtrlUp != null)
                {
                    keyboardHotkeyCtrlUp.Dispose();
                    keyboardHotkeyCtrlUp = null;
                }
                if (keyboardHotkeyCtrlLeft != null)
                {
                    keyboardHotkeyCtrlLeft.Dispose();
                    keyboardHotkeyCtrlLeft = null;
                }
                if (keyboardHotkeyCtrlRight != null)
                {
                    keyboardHotkeyCtrlRight.Dispose();
                    keyboardHotkeyCtrlRight = null;
                }
                if (keyboardHotkeyCtrlDown != null)
                {
                    keyboardHotkeyCtrlDown.Dispose();
                    keyboardHotkeyCtrlDown = null;
                }

                //Save the new position
                programPrefs.loadedData.windowPosX = (int)this.Left;
                programPrefs.loadedData.windowPosY = (int)this.Top;
                programPrefs.Save();

                //Disable the move mode
                isMoveModeEnabled = false;
                //Cancel here
                return;
            }

            //If the move mode is disabled...
            if (isMoveModeEnabled == false)
            {
                //Change the UI to move mode
                bg.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 0, 0));

                //Show a warning
                System.Windows.MessageBox.Show("The \"Move Mode\" is enabled. Use \"Ctrl + Arrow Keys\" to move Resource Metter 2 by 100 pixels in each direction. To move Resource Metter 2 by 1 pixel in each direction, use only the \"Arrow Keys\".", "Move Mode");

                //If the main window is not in topmost, put it in topmost
                if (this.Topmost == false)
                {
                    this.notifyIcon.Icon = new System.Drawing.Icon(@"Content/tray-icon-pinned.ico");
                    this.Topmost = true;
                }

                //Open the position display window
                if (windowPositionDisplay == null)
                {
                    windowPositionDisplay = new WindowPositionDisplay();
                    windowPositionDisplay.Owner = this;
                    windowPositionDisplay.Show();
                    windowPositionDisplay.Visibility = Visibility.Visible;
                    windowPositionDisplay.WindowStartupLocation = WindowStartupLocation.Manual;
                    windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    windowPositionDisplay.Left = ((SystemParameters.PrimaryScreenWidth / 2.0f) - (windowPositionDisplay.Width / 2.0f));
                    windowPositionDisplay.Top = ((SystemParameters.PrimaryScreenHeight) - (windowPositionDisplay.Height) - 48);
                }

                //Set the move mode option checked
                notifyIcon_MoveModeOption.Checked = true;

                //Prepare the key listener
                if (keyboardKeysWatcher == null)
                {
                    keyboardKeysWatcher = new KeyboardKeys_Watcher();
                    keyboardKeysWatcher.OnPressKeys += (int keyCode) =>
                    {
                        //If is the Up arrow...
                        if (((VirtualKeyInt)keyCode) == VirtualKeyInt.VK_UP)
                            this.Top -= 1;
                        //If is the Left arrow...
                        if (((VirtualKeyInt)keyCode) == VirtualKeyInt.VK_LEFT)
                            this.Left -= 1;
                        //If is the Right arrow...
                        if (((VirtualKeyInt)keyCode) == VirtualKeyInt.VK_RIGHT)
                            this.Left += 1;
                        //If is the Down arrow...
                        if (((VirtualKeyInt)keyCode) == VirtualKeyInt.VK_DOWN)
                            this.Top += 1;

                        //Update the position display
                        windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    };
                }
                //Prepare the hotkey ctrl up listener
                if (keyboardHotkeyCtrlUp == null)
                {
                    keyboardHotkeyCtrlUp = new KeyboardHotkey_Interceptor(this, 10, KeyboardHotkey_Interceptor.ModifierKeyCodes.Control, VirtualKeyInt.VK_UP);
                    keyboardHotkeyCtrlUp.OnPressHotkey += () =>
                    {
                        //Move and update position display
                        this.Top -= 99;
                        //Update the position display
                        windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    };
                }
                //Prepare the hotkey ctrl left listener
                if (keyboardHotkeyCtrlLeft == null)
                {
                    keyboardHotkeyCtrlLeft = new KeyboardHotkey_Interceptor(this, 20, KeyboardHotkey_Interceptor.ModifierKeyCodes.Control, VirtualKeyInt.VK_LEFT);
                    keyboardHotkeyCtrlLeft.OnPressHotkey += () =>
                    {
                        //Move and update position display
                        this.Left -= 99;
                        //Update the position display
                        windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    };
                }
                //Prepare the hotkey ctrl right listener
                if (keyboardHotkeyCtrlRight == null)
                {
                    keyboardHotkeyCtrlRight = new KeyboardHotkey_Interceptor(this, 30, KeyboardHotkey_Interceptor.ModifierKeyCodes.Control, VirtualKeyInt.VK_RIGHT);
                    keyboardHotkeyCtrlRight.OnPressHotkey += () =>
                    {
                        //Move and update position display
                        this.Left += 99;
                        //Update the position display
                        windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    };
                }
                //Prepare the hotkey ctrl down listener
                if (keyboardHotkeyCtrlDown == null)
                {
                    keyboardHotkeyCtrlDown = new KeyboardHotkey_Interceptor(this, 40, KeyboardHotkey_Interceptor.ModifierKeyCodes.Control, VirtualKeyInt.VK_DOWN);
                    keyboardHotkeyCtrlDown.OnPressHotkey += () =>
                    {
                        //Move and update position display
                        this.Top += 99;
                        //Update the position display
                        windowPositionDisplay.positionTxt.Text = (this.Left + "x" + this.Top);
                    };
                }

                //Disable all window percent views, that is currently enabled
                if (cpuUsageWindowPercentView != null)
                    ClickToViewCpuUsage();
                if (ramUsageWindowPercentView != null)
                    ClickToViewRamUsage();
                if (gpuUsageWindowPercentView != null)
                    ClickToViewGpuUsage();
                if (vramUsageWindowPercentView != null)
                    ClickToViewVramUsage();
                if (diskUsageWindowPercentView != null)
                    ClickToViewDiskUsage();

                //Enable the move mode
                isMoveModeEnabled = true;
                //Cancel here
                return;
            }
        }

        private void NotifyIcon_WindowsRestorerToggle()
        {
            //If the auto restore windows is enabled...
            if (programPrefs.loadedData.autoRestoreWindows == true)
            {
                //Disable the auto restore windows
                programPrefs.loadedData.autoRestoreWindows = false;
                //Save the preferences
                programPrefs.Save();

                //Disable the check in the option
                notifyIcon_WindowsRestorerOption.Checked = false;

                //Disable the auto restore windows timer
                DisableTheAutoRestoreWindowsTimer();

                //Cancel here
                return;
            }

            //If the auto restore windows is disabled...
            if (programPrefs.loadedData.autoRestoreWindows == false)
            {
                //Show a warning
                System.Windows.MessageBox.Show("The \"Auto Restore Windows\" function has been enabled. Now, all Windows on the Secondary Monitor will be automatically restored if they are minimized.", "Auto Restore Windows");

                //Enable the auto restore windows
                programPrefs.loadedData.autoRestoreWindows = true;
                //Save the preferences
                programPrefs.Save();

                //Enable the check in the option
                notifyIcon_WindowsRestorerOption.Checked = true;

                //Enable the auto restore windows timer
                EnableTheAutoRestoreWindowsTimer();

                //Cancel here
                return;
            }
        }

        private void EnableTheAutoRestoreWindowsTimer()
        {
            //If don't have a timer running...
            if (windowsRestorerTimer == null)
            {
                //Create a new timer
                windowsRestorerTimer = new System.Windows.Forms.Timer { Interval = 3000 };
                windowsRestorerTimer.Enabled = true;
                windowsRestorerTimer.Tick += new EventHandler((object s, EventArgs e) =>
                {
                    //List all current opened windows
                    List<IntPtr> currentOpenedWindowsOnSecondaryMonitor = new List<IntPtr>();

                    //Get all running processes with a interface
                    Process[] processes = Process.GetProcesses();
                    foreach (Process process in processes)
                        if (String.IsNullOrEmpty(process.MainWindowTitle) == false)
                        {
                            //Get all UWP windows opened in secondary monitor
                            if (process.ProcessName == "ApplicationFrameHost")
                                foreach (var handle in EnumerateProcessWindowHandles(process.Id))
                                    if ((GetPlacement(handle).rcNormalPosition.Left + 8) > Screen.PrimaryScreen.Bounds.Width)
                                        currentOpenedWindowsOnSecondaryMonitor.Add(handle);

                            //Get all Win32 windows opened in secondary monitor
                            if (process.ProcessName != "ApplicationFrameHost")
                                if ((GetPlacement(process.MainWindowHandle).rcNormalPosition.Left + 8) > Screen.PrimaryScreen.Bounds.Width)
                                    currentOpenedWindowsOnSecondaryMonitor.Add(process.MainWindowHandle);
                        }

                    //Force to restore all windows on secondary monitor of list
                    foreach (IntPtr windowHandle in currentOpenedWindowsOnSecondaryMonitor)
                        ShowWindow(windowHandle, SW_SHOWNOACTIVATE);
                });
            }
        }

        private void DisableTheAutoRestoreWindowsTimer()
        {
            //If have a timer running...
            if (windowsRestorerTimer != null)
            {
                //Delete the timer
                windowsRestorerTimer.Enabled = false;
                windowsRestorerTimer.Stop();
                windowsRestorerTimer.Dispose();
                windowsRestorerTimer = null;
            }
        }

        private void NotifyIcon_ChangeCpuFanPwmSlot(System.Windows.Forms.ToolStripMenuItem clickedItem)
        {
            //Disable checks of all buttons of CPU Fan
            foreach (System.Windows.Forms.ToolStripMenuItem item in notifyIcon_cpuFanSlotButtons)
                item.Checked = false;

            //Get slot clicked name
            string clickedSlotName = clickedItem.Text;

            //Convert the name to slot index
            int clickedSlot = -1;
            if (clickedSlotName == "Off")
                clickedSlot = -1;
            if (clickedSlotName != "Off")
                clickedSlot = int.Parse(clickedSlotName.Replace("PWM #", ""));

            //Change the slot in preferences
            programPrefs.loadedData.cpuFanPwmSlot = clickedSlot;
            programPrefs.Save();

            //Restore the current selection of PWM slot for CPU Fan
            notifyIcon_cpuFanSlotButtons[(programPrefs.loadedData.cpuFanPwmSlot + 1)].Checked = true;
        }

        private void NotifyIcon_ChangeCpuOptPwmSlot(System.Windows.Forms.ToolStripMenuItem clickedItem)
        {
            //Disable checks of all buttons of CPU Opt
            foreach (System.Windows.Forms.ToolStripMenuItem item in notifyIcon_cpuOptSlotButtons)
                item.Checked = false;

            //Get slot clicked name
            string clickedSlotName = clickedItem.Text;

            //Convert the name to slot index
            int clickedSlot = -1;
            if (clickedSlotName == "Off")
                clickedSlot = -1;
            if (clickedSlotName != "Off")
                clickedSlot = int.Parse(clickedSlotName.Replace("PWM #", ""));

            //Change the slot in preferences
            programPrefs.loadedData.cpuOptPwmSlot = clickedSlot;
            programPrefs.Save();

            //Restore the current selection of PWM slot for CPU Opt
            notifyIcon_cpuOptSlotButtons[(programPrefs.loadedData.cpuOptPwmSlot + 1)].Checked = true;
        }

        private void NotifyIcon_ChangeCpuPumpPwmSlot(System.Windows.Forms.ToolStripMenuItem clickedItem)
        {
            //Disable checks of all buttons of CPU Pump
            foreach (System.Windows.Forms.ToolStripMenuItem item in notifyIcon_cpuPumpSlotButtons)
                item.Checked = false;

            //Get slot clicked name
            string clickedSlotName = clickedItem.Text;

            //Convert the name to slot index
            int clickedSlot = -1;
            if (clickedSlotName == "Off")
                clickedSlot = -1;
            if (clickedSlotName != "Off")
                clickedSlot = int.Parse(clickedSlotName.Replace("PWM #", ""));

            //Change the slot in preferences
            programPrefs.loadedData.cpuPumpPwmSlot = clickedSlot;
            programPrefs.Save();

            //Restore the current selection of PWM slot for CPU Pump
            notifyIcon_cpuPumpSlotButtons[(programPrefs.loadedData.cpuPumpPwmSlot + 1)].Checked = true;
        }

        private void NotifyIcon_ChangePumpRpmWarn(int newWarnValue)
        {
            //Disable checks of all buttons of Pump RPM Warn
            foreach (var item in notifyIcon_pumpRpmWarnButtons)
                item.Value.Checked = false;

            //Change the pump rpm warn in preferences
            programPrefs.loadedData.cpuPumpRpmWarn = newWarnValue;
            programPrefs.Save();

            //Restore the current selection of Pump RPM Warn
            notifyIcon_pumpRpmWarnButtons[programPrefs.loadedData.cpuPumpRpmWarn.ToString()].Checked = true;
        }

        private void NotifyIcon_StartWithSystemToggle()
        {
            //If the auto start with windows is enabled...
            if (programPrefs.loadedData.autoStartWithSystem == true)
            {
                //Enable the auto start with windows
                programPrefs.loadedData.autoStartWithSystem = false;
                //Save the preferences
                programPrefs.Save();

                //Enable the check in the option
                notifyIcon_AutoStartOption.Checked = false;

                //Unregister this program from boot
                StartupManager startupManager = new StartupManager("Resources Metter 2");
                bool success = startupManager.RegisterToStartOnBoot(false);
                if (success == false)
                    System.Windows.MessageBox.Show("Unable to apply the Boot settings correctly.", "Error");

                //Cancel here
                return;
            }

            //If the auto start with windows is disabled...
            if (programPrefs.loadedData.autoStartWithSystem == false)
            {
                //Enable the auto start with windows
                programPrefs.loadedData.autoStartWithSystem = true;
                //Save the preferences
                programPrefs.Save();

                //Enable the check in the option
                notifyIcon_AutoStartOption.Checked = true;

                //Register this program from boot
                StartupManager startupManager = new StartupManager("Resources Metter 2");
                bool success = startupManager.RegisterToStartOnBoot(true);
                if (success == false)
                    System.Windows.MessageBox.Show("Unable to apply the Boot settings correctly.", "Error");

                //Cancel here
                return;
            }
        }

        private void NotifyIcon_LibreHardwareMonitorOpen()
        {
            //Start the routine to start and manage the Libre Hardware Monitor
            IDisposable lhmStartRoutine = Coroutine.Start(LibreHardwareMonitorStartRoutine());
        }

        private IEnumerator LibreHardwareMonitorStartRoutine()
        {
            //Disable the button
            notifyIcon_LibreHMOption.Enabled = false;

            //Start the process of Libre Hardware Monitor
            libreHardwareMonitorProcess = new Process();
            libreHardwareMonitorProcess.StartInfo.FileName = System.IO.Path.Combine((Directory.GetCurrentDirectory() + @"\Content\LHM"), "LibreHardwareMonitor.exe");
            libreHardwareMonitorProcess.StartInfo.WorkingDirectory = (Directory.GetCurrentDirectory() + @"\Content\LHM");
            libreHardwareMonitorProcess.Start();

            //Create the loop interval
            WaitForSeconds loopInterval = new WaitForSeconds(0.5f);

            //Wait until the process finish
            while (true)
            {
                //If was finishes, break the loop
                if (libreHardwareMonitorProcess.HasExited == true)
                    break;

                //Wait the interval
                yield return loopInterval;
            }

            //Remove the reference for the process
            libreHardwareMonitorProcess = null;

            //Enable the button again
            notifyIcon_LibreHMOption.Enabled = true;
        }

        private void ClickToViewCpuUsage()
        {
            //If the window is not open, open it
            if (cpuUsageWindowPercentView == null)
            {
                //Elemento to use as reference
                FrameworkElement refControl = cpuUsageProgressBar;

                //Open the window
                cpuUsageWindowPercentView = new WindowPercentView();
                cpuUsageWindowPercentView.Owner = this;
                cpuUsageWindowPercentView.Show();
                cpuUsageWindowPercentView.Visibility = Visibility.Visible;
                cpuUsageWindowPercentView.WindowStartupLocation = WindowStartupLocation.Manual;
                cpuUsageWindowPercentView.Left = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).X + 7.0f) - (cpuUsageWindowPercentView.Width / 2.0f));
                cpuUsageWindowPercentView.Top = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).Y - 8.0f) - (cpuUsageWindowPercentView.Height));

                //Update the UI of the window
                cpuUsageWindowPercentView.icon.Source = new BitmapImage(new Uri(@"pack://application:,,,/Resources/cpu-icon.png"));
                cpuUsageWindowPercentView.value.Text = "-";

                //Cancel here
                return;
            }

            //If the window is already open, close it
            if (cpuUsageWindowPercentView != null)
            {
                //Close the window
                cpuUsageWindowPercentView.Close();
                cpuUsageWindowPercentView = null;

                //Cancel here
                return;
            }
        }

        private void ClickToViewRamUsage()
        {
            //If the window is not open, open it
            if (ramUsageWindowPercentView == null)
            {
                //Elemento to use as reference
                FrameworkElement refControl = ramUsageProgressBar;

                //Open the window
                ramUsageWindowPercentView = new WindowPercentView();
                ramUsageWindowPercentView.Owner = this;
                ramUsageWindowPercentView.Show();
                ramUsageWindowPercentView.Visibility = Visibility.Visible;
                ramUsageWindowPercentView.WindowStartupLocation = WindowStartupLocation.Manual;
                ramUsageWindowPercentView.Left = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).X + 7.0f) - (ramUsageWindowPercentView.Width / 2.0f));
                ramUsageWindowPercentView.Top = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).Y - 8.0f) - (ramUsageWindowPercentView.Height));

                //Update the UI of the window
                ramUsageWindowPercentView.icon.Source = new BitmapImage(new Uri(@"pack://application:,,,/Resources/ram-icon.png"));
                ramUsageWindowPercentView.icon.LayoutTransform = new RotateTransform(180);
                ramUsageWindowPercentView.value.Text = "-";

                //Cancel here
                return;
            }

            //If the window is already open, close it
            if (ramUsageWindowPercentView != null)
            {
                //Close the window
                ramUsageWindowPercentView.Close();
                ramUsageWindowPercentView = null;

                //Cancel here
                return;
            }
        }

        private void ClickToViewGpuUsage()
        {
            //If the window is not open, open it
            if (gpuUsageWindowPercentView == null)
            {
                //Elemento to use as reference
                FrameworkElement refControl = gpuUsageProgressBar;

                //Open the window
                gpuUsageWindowPercentView = new WindowPercentView();
                gpuUsageWindowPercentView.Owner = this;
                gpuUsageWindowPercentView.Show();
                gpuUsageWindowPercentView.Visibility = Visibility.Visible;
                gpuUsageWindowPercentView.WindowStartupLocation = WindowStartupLocation.Manual;
                gpuUsageWindowPercentView.Left = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).X + 7.0f) - (gpuUsageWindowPercentView.Width / 2.0f));
                gpuUsageWindowPercentView.Top = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).Y - 8.0f) - (gpuUsageWindowPercentView.Height));

                //Update the UI of the window
                gpuUsageWindowPercentView.icon.Source = new BitmapImage(new Uri(@"pack://application:,,,/Resources/gpu-card-icon2.png"));
                gpuUsageWindowPercentView.value.Text = "-";

                //Cancel here
                return;
            }

            //If the window is already open, close it
            if (gpuUsageWindowPercentView != null)
            {
                //Close the window
                gpuUsageWindowPercentView.Close();
                gpuUsageWindowPercentView = null;

                //Cancel here
                return;
            }
        }

        private void ClickToViewVramUsage()
        {
            //If the window is not open, open it
            if (vramUsageWindowPercentView == null)
            {
                //Elemento to use as reference
                FrameworkElement refControl = vramUsageProgressBar;

                //Open the window
                vramUsageWindowPercentView = new WindowPercentView();
                vramUsageWindowPercentView.Owner = this;
                vramUsageWindowPercentView.Show();
                vramUsageWindowPercentView.Visibility = Visibility.Visible;
                vramUsageWindowPercentView.WindowStartupLocation = WindowStartupLocation.Manual;
                vramUsageWindowPercentView.Left = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).X + 7.0f) - (vramUsageWindowPercentView.Width / 2.0f));
                vramUsageWindowPercentView.Top = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).Y - 8.0f) - (vramUsageWindowPercentView.Height));

                //Update the UI of the window
                vramUsageWindowPercentView.icon.Source = new BitmapImage(new Uri(@"pack://application:,,,/Resources/gpu-ram-icon.png"));
                vramUsageWindowPercentView.value.Text = "-";

                //Cancel here
                return;
            }

            //If the window is already open, close it
            if (vramUsageWindowPercentView != null)
            {
                //Close the window
                vramUsageWindowPercentView.Close();
                vramUsageWindowPercentView = null;

                //Cancel here
                return;
            }
        }

        private void ClickToViewDiskUsage()
        {
            //If the window is not open, open it
            if (diskUsageWindowPercentView == null)
            {
                //Elemento to use as reference
                FrameworkElement refControl = diskUsageProgressBar;

                //Open the window
                diskUsageWindowPercentView = new WindowPercentView();
                diskUsageWindowPercentView.Owner = this;
                diskUsageWindowPercentView.Show();
                diskUsageWindowPercentView.Visibility = Visibility.Visible;
                diskUsageWindowPercentView.WindowStartupLocation = WindowStartupLocation.Manual;
                diskUsageWindowPercentView.Left = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).X + 7.0f) - (diskUsageWindowPercentView.Width / 2.0f));
                diskUsageWindowPercentView.Top = ((refControl.PointToScreen(new System.Windows.Point(0.0d, 0.0d)).Y - 8.0f) - (diskUsageWindowPercentView.Height));

                //Update the UI of the window
                diskUsageWindowPercentView.icon.Source = new BitmapImage(new Uri(@"pack://application:,,,/Resources/hdd-icon-black.png"));
                diskUsageWindowPercentView.value.Text = "-";

                //Cancel here
                return;
            }

            //If the window is already open, close it
            if (diskUsageWindowPercentView != null)
            {
                //Close the window
                diskUsageWindowPercentView.Close();
                diskUsageWindowPercentView = null;

                //Cancel here
                return;
            }
        }

        private void QuitApp()
        {
            //If the Libre Hardware Monitor is running, stop it
            if (libreHardwareMonitorProcess != null)
                libreHardwareMonitorProcess.Kill();

            //Quit from the app
            System.Windows.Application.Current.Shutdown();
        }

        //Private auxiliar APIs

        #region StartUpManagerUsingDLLTaskManager
        public class StartupManager
        {
            //Private variable
            public string programName = "Resources Metter 2";

            //Public variables
            public bool taskSchedulerAvailable = false;

            //Core methods
            public StartupManager(string programName)
            {
                //Fill needed variables
                this.programName = programName;

                //Check if task scheduler instance is connected
                bool isTaskSchedulerConnected = Microsoft.Win32.TaskScheduler.TaskService.Instance.Connected;

                //If task scheduler is available
                if (isAdministrator() == true && isTaskSchedulerConnected == true)
                    taskSchedulerAvailable = true;
                //If task scheduler is not available
                if (isAdministrator() == false || isTaskSchedulerConnected == false)
                    taskSchedulerAvailable = false;
            }

            public bool RegisterToStartOnBoot(bool startOnBoot)
            {
                //Try to register
                try
                {
                    //If is desired to register and task scheduler is available
                    if (startOnBoot == true && taskSchedulerAvailable == true)
                        if (isOnTaskScheduler() == false) //<- Register on task scheduler (if is not on the task scheduler)
                        {
                            //Create the task scheduling
                            Microsoft.Win32.TaskScheduler.TaskDefinition taskDefinition = Microsoft.Win32.TaskScheduler.TaskService.Instance.NewTask();
                            taskDefinition.RegistrationInfo.Description = "Starts Resources Metter 2 on Windows startup.";
                            //Set the trigger
                            taskDefinition.Triggers.Add(new Microsoft.Win32.TaskScheduler.LogonTrigger());
                            //Set the preferences
                            taskDefinition.Settings.StartWhenAvailable = true;
                            taskDefinition.Settings.DisallowStartIfOnBatteries = false;
                            taskDefinition.Settings.StopIfGoingOnBatteries = false;
                            taskDefinition.Settings.ExecutionTimeLimit = TimeSpan.Zero;
                            taskDefinition.Settings.AllowHardTerminate = false;
                            //Set the levels
                            taskDefinition.Principal.RunLevel = Microsoft.Win32.TaskScheduler.TaskRunLevel.Highest;
                            taskDefinition.Principal.LogonType = Microsoft.Win32.TaskScheduler.TaskLogonType.InteractiveToken;
                            //Set the actions to run
                            taskDefinition.Actions.Add(new Microsoft.Win32.TaskScheduler.ExecAction((Directory.GetCurrentDirectory() + @"\Resources Metter 2.exe"), "", Directory.GetCurrentDirectory()));
                            //Register the task to be runned
                            Microsoft.Win32.TaskScheduler.TaskService.Instance.RootFolder.RegisterTaskDefinition(nameof(Resources_Metter_2), taskDefinition);
                        }
                    //If is desired to register and task scheduler is not available
                    if (startOnBoot == true && taskSchedulerAvailable == false)
                        SetRegisterInStartup(true); //<- Add to registry
                    //If is desired to unregister
                    if (startOnBoot == false)
                    {
                        //Remove from task scheduler (if is on task scheduler)
                        if (taskSchedulerAvailable == true && isOnTaskScheduler() == true)
                        {
                            Microsoft.Win32.TaskScheduler.Task task = Microsoft.Win32.TaskScheduler.TaskService.Instance.AllTasks.FirstOrDefault(x => x.Name.Equals(nameof(Resources_Metter_2), StringComparison.OrdinalIgnoreCase));
                            if (task != null)
                                task.Folder.DeleteTask(task.Name, false);
                        }
                        //Remove from registry
                        SetRegisterInStartup(false);
                    }

                    //Return success
                    return true;
                }
                catch (Exception ex)
                {
                    //Show log error
                    System.IO.File.WriteAllText("last-startupmanager-error.txt", ex.Message);
                    //Return error
                    return false;
                }
            }

            //Tools methods

            private bool isAdministrator()
            {
                try
                {
                    System.Security.Principal.WindowsIdentity identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                    System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(identity);

                    return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
                }
                catch
                {
                    return false;
                }
            }

            private bool isOnTaskScheduler()
            {
                //Check if start on boot task is scheduled
                Microsoft.Win32.TaskScheduler.Task task = Microsoft.Win32.TaskScheduler.TaskService.Instance.AllTasks.FirstOrDefault(x => x.Name.Equals(nameof(Resources_Metter_2), StringComparison.OrdinalIgnoreCase));
                //If task object is null, cancel
                if (task == null)
                    return false;

                //Try to find a task scheduled to run this app
                foreach (Microsoft.Win32.TaskScheduler.Action action in task.Definition.Actions)
                    if (action.ActionType == Microsoft.Win32.TaskScheduler.TaskActionType.Execute && action is Microsoft.Win32.TaskScheduler.ExecAction execAction)
                        if (execAction.Path.Equals(System.Reflection.Assembly.GetExecutingAssembly().Location, StringComparison.OrdinalIgnoreCase))
                            return true;

                //Default return
                return false;
            }

            //Fallback method
            private void SetRegisterInStartup(bool startOnBoot)
            {
                //Register or unregister this app from boot
                Microsoft.Win32.RegistryKey registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                if (startOnBoot == true)
                {
                    registryKey.SetValue(programName, "\"" + (Directory.GetCurrentDirectory() + @"\Resources Metter 2.exe") + "\"");
                    registryKey.Close();
                }
                if (startOnBoot == false && registryKey.GetValue(programName) != null)
                {
                    registryKey.DeleteValue(programName);
                    registryKey.Close();
                }
            }
        }
        #endregion

        #region LibreHardwareMonitorUpdateVisitor
        public class UpdateVisitor : IVisitor
        {
            public void VisitComputer(IComputer computer)
            {
                computer.Traverse(this);
            }
            public void VisitHardware(IHardware hardware)
            {
                hardware.Update();
                foreach (IHardware subHardware in hardware.SubHardware) subHardware.Accept(this);
            }
            public void VisitSensor(ISensor sensor) { }
            public void VisitParameter(IParameter parameter) { }
        }
        #endregion

        #region DetectionAndRestorationOfAllWindowsOpenedOnSecondaryMonitor

        //Dependencies of DLLs
        private int SW_SHOWNOACTIVATE = 4;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);
        delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        private IEnumerable<IntPtr> EnumerateProcessWindowHandles(int processId)
        {
            var handles = new List<IntPtr>();

            foreach (ProcessThread thread in Process.GetProcessById(processId).Threads)
                EnumThreadWindows(thread.Id, (hWnd, lParam) => { handles.Add(hWnd); return true; }, IntPtr.Zero);

            return handles;
        }

        private static WINDOWPLACEMENT GetPlacement(IntPtr hwnd)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(hwnd, ref placement);
            return placement;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        internal enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }

        #endregion

        #region KeyboardWatcher

        public enum VirtualKeyHex
        {
            VK_ESCAPE = 0x1B,            //<- ESC
            VK_TAB = 0x09,
            VK_SHIFT = 0x10,
            VK_LSHIFT = 0xA0,
            VK_RSHIFT = 0xA1,
            VK_CONTROL = 0x11,
            VK_LCONTROL = 0xA2,
            VK_RCONTROL = 0xA3,
            VK_MENU = 0x12,              //<- ALT
            VK_LMENU = 0xA4,             //<- LEFT_ALT
            VK_RMENU = 0xA5,             //<- RIGHT_ALT
            VK_CAPITAL = 0x14,           //<- CAPS_LOCK
            VK_NUMLOCK = 0x90,
            VK_SCROLL = 0x91,            //<- SCROLL_LOCK
            VK_RETURN = 0x0D,            //<- ENTER
            VK_SPACE = 0x20,
            VK_PRIOR = 0x21,             //<- PAGE_UP
            VK_NEXT = 0x22,              //<- PAGE_DOWN
            VK_END = 0x23,
            VK_HOME = 0x24,
            VK_LEFT = 0x25,              //<- LEFT_ARROW
            VK_UP = 0x26,                //<- UP_ARROW
            VK_RIGHT = 0x27,             //<- RIGHT_ARROW
            VK_DOWN = 0x28,              //<- DOWN_ARROW
            VK_SNAPSHOT = 0x2C,          //<- PRINT_SCREEN
            VK_PAUSE = 0x13,
            VK_INSERT = 0x2D,
            VK_DELETE = 0x2E,
            VK_BACK = 0x08,              //<- BACKSPACE
            VK_LWIN = 0x5B,              //<- LEFT_WIN
            VK_RWIN = 0x5C,              //<- RIGHT_WIN
            VK_APPS = 0x5D,              //<- APPS_KEY/CONTEXT_MENU
            VK_NUMPAD0 = 0x60,           //<- If NumPad ON: "0" - If NumPad OFF: "INSERT"
            VK_NUMPAD1 = 0x61,           //<- If NumPad ON: "1" - If NumPad OFF: "END"
            VK_NUMPAD2 = 0x62,           //<- If NumPad ON: "2" - If NumPad OFF: "DOWN_ARROW"
            VK_NUMPAD3 = 0x63,           //<- If NumPad ON: "3" - If NumPad OFF: "PAGE_DOWN"
            VK_NUMPAD4 = 0x64,           //<- If NumPad ON: "4" - If NumPad OFF: "LEFT_ARROW"
            VK_NUMPAD5 = 0x65,           //<- If NumPad ON: "5" - If NumPad OFF: "?"
            VK_NUMPAD6 = 0x66,           //<- If NumPad ON: "6" - If NumPad OFF: "RIGHT_ARROW"
            VK_NUMPAD7 = 0x67,           //<- If NumPad ON: "7" - If NumPad OFF: "HOME"
            VK_NUMPAD8 = 0x68,           //<- If NumPad ON: "8" - If NumPad OFF: "UP_ARROW"
            VK_NUMPAD9 = 0x69,           //<- If NumPad ON: "9" - If NumPad OFF: "PAGE_UP"
            VK_MULTIPLY = 0x6A,          //<- NUMPAD_* Can be used with NumPad ON or OFF
            VK_ADD = 0x6B,               //<- NUMPAD_+ Can be used with NumPad ON or OFF
            VK_SEPARATOR = 0x6C,         //<- NUMPAD_. Can be used with NumPad ON or OFF
            VK_SUBTRACT = 0x6D,          //<- NUMPAD_- Can be used with NumPad ON or OFF
            VK_DECIMAL = 0x6E,           //<- NUMPAD_, If NumPad ON: "," - If NumPad OFF: "DELETE"
            VK_DIVIDE = 0x6F,            //<- NUMPAD_/ Can be used with NumPad ON or OFF
            VK_F1 = 0x70,
            VK_F2 = 0x71,
            VK_F3 = 0x72,
            VK_F4 = 0x73,
            VK_F5 = 0x74,
            VK_F6 = 0x75,
            VK_F7 = 0x76,
            VK_F8 = 0x77,
            VK_F9 = 0x78,
            VK_F10 = 0x79,
            VK_F11 = 0x7A,
            VK_F12 = 0x7B,
            VK_0 = 0x30,
            VK_1 = 0x31,
            VK_2 = 0x32,
            VK_3 = 0x33,
            VK_4 = 0x34,
            VK_5 = 0x35,
            VK_6 = 0x36,
            VK_7 = 0x37,
            VK_8 = 0x38,
            VK_9 = 0x39,
            VK_A = 0x41,
            VK_B = 0x42,
            VK_C = 0x43,
            VK_D = 0x44,
            VK_E = 0x45,
            VK_F = 0x46,
            VK_G = 0x47,
            VK_H = 0x48,
            VK_I = 0x49,
            VK_J = 0x4A,
            VK_K = 0x4B,
            VK_L = 0x4C,
            VK_M = 0x4D,
            VK_N = 0x4E,
            VK_O = 0x4F,
            VK_P = 0x50,
            VK_Q = 0x51,
            VK_R = 0x52,
            VK_S = 0x53,
            VK_T = 0x54,
            VK_U = 0x55,
            VK_V = 0x56,
            VK_W = 0x57,
            VK_X = 0x58,
            VK_Y = 0x59,
            VK_Z = 0x5A,
            VK_OEM_1 = 0xBA,             //<- Vary according to Keyboard Layout. US: ";"      - BR: "ç"
            VK_OEM_PLUS = 0xBB,          //<- Same for ALL Keyboard Layout.     ALL: "= or +"
            VK_OEM_COMMA = 0xBC,         //<- Same for ALL Keyboard Layout.     ALL: ", or <"
            VK_OEM_MINUS = 0xBD,         //<- Same for ALL Keyboard Layout.     ALL: "- or _"
            VK_OEM_PERIOD = 0xBE,        //<- Same for ALL Keyboard Layout.     ALL: ". or >"
            VK_OEM_2 = 0xBF,             //<- Vary according to Keyboard Layout. US: "/ or ?" - BR: "; or :"
            VK_OEM_3 = 0xC0,             //<- Vary according to Keyboard Layout. US: "` or ~" - BR: "' or ""
            VK_OEM_4 = 0xDB,             //<- Vary according to Keyboard Layout. US: "[ or {" - BR: "´ or `"
            VK_OEM_5 = 0xDC,             //<- Vary according to Keyboard Layout. US: "\ or |" - BR: "] or }"
            VK_OEM_6 = 0xDD,             //<- Vary according to Keyboard Layout. US: "] or }" - BR: "[ or {"
            VK_OEM_7 = 0xDE,             //<- Vary according to Keyboard Layout. US: "' or "" - BR: "~ or ^"
            VK_OEM_8 = 0xDF,             //<- Vary according to Keyboard Layout. CA: "LEFT_CONTROL"
            VK_OEM_102 = 0xE2,           //<- Vary according to Keyboard Layout. EU: "\ or |" - BR: "\ or |"
            VK_PACKET = 0xE7,            //<- Used to pass Unicode characters as if they were keystrokes.
            VK_ATTN = 0xF6,              //<- ATTN_KEY       (Attention Key in the context of mainframe)
            VK_CRSEL = 0xF7,             //<- CRSEL_KEY      (Cursor Select Key)
            VK_EXSEL = 0xF8,             //<- EXSEL_KEY      (Extend Selection Key)
            VK_EREOF = 0xF9,             //<- ERASE_EOF_KEY  (Delete all characters from the current cursor position to the end of the field or line, in the context of mainframe)
            VK_PLAY = 0xFA,              //<- PLAY_KEY       (Alternative to VK_MEDIA_PLAY_PAUSE)
            VK_ZOOM = 0xFB,              //<- ZOOM_KEY       (Refers to a dedicated key found on certain specialized keyboards, most notably the Microsoft Natural Ergonomic Keyboard 4000)
            VK_NONAME = 0xFC,            //<- NO_NAME_KEY    (A blank-key of keyboard, which has no lettering, or it could be a generic, unbranded keyboard)
            VK_PA1 = 0xFD,               //<- PA1_KEY        (A special function key on legacy keyboards, especially used in IBM mainframes, and can also refer to a musical keyboard model)
            VK_BROWSER_BACK = 0xA6,      //<- MULTIMEDIA_KEY
            VK_BROWSER_FORWARD = 0xA7,   //<- MULTIMEDIA_KEY
            VK_BROWSER_REFRESH = 0xA8,   //<- MULTIMEDIA_KEY
            VK_BROWSER_STOP = 0xA9,      //<- MULTIMEDIA_KEY
            VK_BROWSER_SEARCH = 0xAA,    //<- MULTIMEDIA_KEY
            VK_BROWSER_FAVORITES = 0xAB, //<- MULTIMEDIA_KEY
            VK_BROWSER_HOME = 0xAC,      //<- MULTIMEDIA_KEY
            VK_VOLUME_MUTE = 0xAD,       //<- MULTIMEDIA_KEY
            VK_VOLUME_DOWN = 0xAE,       //<- MULTIMEDIA_KEY
            VK_VOLUME_UP = 0xAF,         //<- MULTIMEDIA_KEY
            VK_MEDIA_NEXT_TRACK = 0xB0,  //<- MULTIMEDIA_KEY
            VK_MEDIA_PREV_TRACK = 0xB1,  //<- MULTIMEDIA_KEY
            VK_MEDIA_STOP = 0xB2,        //<- MULTIMEDIA_KEY
            VK_MEDIA_PLAY_PAUSE = 0xB3,  //<- MULTIMEDIA_KEY (Alternative to PLAY_KEY)
            VK_LAUNCH_MAIL = 0xB4,       //<- MULTIMEDIA_KEY
        }

        public enum VirtualKeyInt
        {
            VK_ESCAPE = 27,              //<- ESC
            VK_TAB = 9,
            VK_SHIFT = 16,
            VK_LSHIFT = 160,
            VK_RSHIFT = 161,
            VK_CONTROL = 17,
            VK_LCONTROL = 162,
            VK_RCONTROL = 163,
            VK_MENU = 18,                //<- ALT
            VK_LMENU = 164,              //<- LEFT_ALT
            VK_RMENU = 165,              //<- RIGHT_ALT
            VK_CAPITAL = 20,             //<- CAPS_LOCK
            VK_NUMLOCK = 144,
            VK_SCROLL = 145,             //<- SCROLL_LOCK
            VK_RETURN = 13,              //<- ENTER
            VK_SPACE = 32,
            VK_PRIOR = 33,               //<- PAGE_UP
            VK_NEXT = 34,                //<- PAGE_DOWN
            VK_END = 35,
            VK_HOME = 36,
            VK_LEFT = 37,                //<- LEFT_ARROW
            VK_UP = 38,                  //<- UP_ARROW
            VK_RIGHT = 39,               //<- RIGHT_ARROW
            VK_DOWN = 40,                //<- DOWN_ARROW
            VK_SNAPSHOT = 44,            //<- PRINT_SCREEN
            VK_PAUSE = 19,
            VK_INSERT = 45,
            VK_DELETE = 46,
            VK_BACK = 8,                 //<- BACKSPACE
            VK_LWIN = 91,                //<- LEFT_WIN
            VK_RWIN = 92,                //<- RIGHT_WIN
            VK_APPS = 93,                //<- APPS_KEY/CONTEXT_MENU
            VK_NUMPAD0 = 96,             //<- If NumPad ON: "0" - If NumPad OFF: "INSERT"
            VK_NUMPAD1 = 97,             //<- If NumPad ON: "1" - If NumPad OFF: "END"
            VK_NUMPAD2 = 98,             //<- If NumPad ON: "2" - If NumPad OFF: "DOWN_ARROW"
            VK_NUMPAD3 = 99,             //<- If NumPad ON: "3" - If NumPad OFF: "PAGE_DOWN"
            VK_NUMPAD4 = 100,            //<- If NumPad ON: "4" - If NumPad OFF: "LEFT_ARROW"
            VK_NUMPAD5 = 101,            //<- If NumPad ON: "5" - If NumPad OFF: "?"
            VK_NUMPAD6 = 102,            //<- If NumPad ON: "6" - If NumPad OFF: "RIGHT_ARROW"
            VK_NUMPAD7 = 103,            //<- If NumPad ON: "7" - If NumPad OFF: "HOME"
            VK_NUMPAD8 = 104,            //<- If NumPad ON: "8" - If NumPad OFF: "UP_ARROW"
            VK_NUMPAD9 = 105,            //<- If NumPad ON: "9" - If NumPad OFF: "PAGE_UP"
            VK_MULTIPLY = 106,           //<- NUMPAD_* Can be used with NumPad ON or OFF
            VK_ADD = 107,                //<- NUMPAD_+ Can be used with NumPad ON or OFF
            VK_SEPARATOR = 108,          //<- NUMPAD_. Can be used with NumPad ON or OFF
            VK_SUBTRACT = 109,           //<- NUMPAD_- Can be used with NumPad ON or OFF
            VK_DECIMAL = 110,            //<- NUMPAD_, If NumPad ON: "," - If NumPad OFF: "DELETE"
            VK_DIVIDE = 111,             //<- NUMPAD_/ Can be used with NumPad ON or OFF
            VK_F1 = 112,
            VK_F2 = 113,
            VK_F3 = 114,
            VK_F4 = 115,
            VK_F5 = 116,
            VK_F6 = 117,
            VK_F7 = 118,
            VK_F8 = 119,
            VK_F9 = 120,
            VK_F10 = 121,
            VK_F11 = 122,
            VK_F12 = 123,
            VK_0 = 48,
            VK_1 = 49,
            VK_2 = 50,
            VK_3 = 51,
            VK_4 = 52,
            VK_5 = 53,
            VK_6 = 54,
            VK_7 = 55,
            VK_8 = 56,
            VK_9 = 57,
            VK_A = 65,
            VK_B = 66,
            VK_C = 67,
            VK_D = 68,
            VK_E = 69,
            VK_F = 70,
            VK_G = 71,
            VK_H = 72,
            VK_I = 73,
            VK_J = 74,
            VK_K = 75,
            VK_L = 76,
            VK_M = 77,
            VK_N = 78,
            VK_O = 79,
            VK_P = 80,
            VK_Q = 81,
            VK_R = 82,
            VK_S = 83,
            VK_T = 84,
            VK_U = 85,
            VK_V = 86,
            VK_W = 87,
            VK_X = 88,
            VK_Y = 89,
            VK_Z = 90,
            VK_OEM_1 = 186,              //<- Vary according to Keyboard Layout. US: ";"      - BR: "ç"
            VK_OEM_PLUS = 187,           //<- Same for ALL Keyboard Layout.     ALL: "= or +"
            VK_OEM_COMMA = 188,          //<- Same for ALL Keyboard Layout.     ALL: ", or <"
            VK_OEM_MINUS = 189,          //<- Same for ALL Keyboard Layout.     ALL: "- or _"
            VK_OEM_PERIOD = 190,         //<- Same for ALL Keyboard Layout.     ALL: ". or >"
            VK_OEM_2 = 191,              //<- Vary according to Keyboard Layout. US: "/ or ?" - BR: "; or :"
            VK_OEM_3 = 192,              //<- Vary according to Keyboard Layout. US: "` or ~" - BR: "' or ""
            VK_OEM_4 = 219,              //<- Vary according to Keyboard Layout. US: "[ or {" - BR: "´ or `"
            VK_OEM_5 = 220,              //<- Vary according to Keyboard Layout. US: "\ or |" - BR: "] or }"
            VK_OEM_6 = 221,              //<- Vary according to Keyboard Layout. US: "] or }" - BR: "[ or {"
            VK_OEM_7 = 222,              //<- Vary according to Keyboard Layout. US: "' or "" - BR: "~ or ^"
            VK_OEM_8 = 223,              //<- Vary according to Keyboard Layout. CA: "LEFT_CONTROL"
            VK_OEM_102 = 226,            //<- Vary according to Keyboard Layout. EU: "\ or |" - BR: "\ or |"
            VK_PACKET = 231,             //<- Used to pass Unicode characters as if they were keystrokes.
            VK_ATTN = 246,               //<- ATTN_KEY       (Attention Key in the context of mainframe)
            VK_CRSEL = 247,              //<- CRSEL_KEY      (Cursor Select Key)
            VK_EXSEL = 248,              //<- EXSEL_KEY      (Extend Selection Key)
            VK_EREOF = 249,              //<- ERASE_EOF_KEY  (Delete all characters from the current cursor position to the end of the field or line, in the context of mainframe)
            VK_PLAY = 250,               //<- PLAY_KEY       (Alternative to VK_MEDIA_PLAY_PAUSE)
            VK_ZOOM = 251,               //<- ZOOM_KEY       (Refers to a dedicated key found on certain specialized keyboards, most notably the Microsoft Natural Ergonomic Keyboard 4000)
            VK_NONAME = 252,             //<- NO_NAME_KEY    (A blank-key of keyboard, which has no lettering, or it could be a generic, unbranded keyboard)
            VK_PA1 = 253,                //<- PA1_KEY        (A special function key on legacy keyboards, especially used in IBM mainframes, and can also refer to a musical keyboard model)
            VK_BROWSER_BACK = 166,       //<- MULTIMEDIA_KEY
            VK_BROWSER_FORWARD = 167,    //<- MULTIMEDIA_KEY
            VK_BROWSER_REFRESH = 168,    //<- MULTIMEDIA_KEY
            VK_BROWSER_STOP = 169,       //<- MULTIMEDIA_KEY
            VK_BROWSER_SEARCH = 170,     //<- MULTIMEDIA_KEY
            VK_BROWSER_FAVORITES = 171,  //<- MULTIMEDIA_KEY
            VK_BROWSER_HOME = 172,       //<- MULTIMEDIA_KEY
            VK_VOLUME_MUTE = 173,        //<- MULTIMEDIA_KEY
            VK_VOLUME_DOWN = 174,        //<- MULTIMEDIA_KEY
            VK_VOLUME_UP = 175,          //<- MULTIMEDIA_KEY
            VK_MEDIA_NEXT_TRACK = 176,   //<- MULTIMEDIA_KEY
            VK_MEDIA_PREV_TRACK = 177,   //<- MULTIMEDIA_KEY
            VK_MEDIA_STOP = 178,         //<- MULTIMEDIA_KEY
            VK_MEDIA_PLAY_PAUSE = 179,   //<- MULTIMEDIA_KEY (Alternative to PLAY_KEY)
            VK_LAUNCH_MAIL = 180,        //<- MULTIMEDIA_KEY
        }

        public class KeyboardKeys_Watcher
        {
            //Private constants
            private const int WH_KEYBOARD_LL = 13;
            private const int WM_KEYDOWN = 0x0100;
            private const int WM_KEYUP = 0x0101;

            //Private variables
            private LowLevelKeyboardProc procedure = null;
            private IntPtr hookId = IntPtr.Zero;
            private bool isDisposed = false;

            //Public callbacks

            public event Action<int> OnPressKeys;

            //Import methods

            private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);

            //Core methods

            public KeyboardKeys_Watcher()
            {
                //Store the callback into a strong reference
                procedure = HookCallback;

                //Register the callback to the low level keys, of windows hook
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    hookId = SetWindowsHookEx(WH_KEYBOARD_LL, procedure, GetModuleHandle(curModule.ModuleName), 0);
                }
            }

            private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
            {
                //If the key code is major than zero, continues to process
                if (nCode >= 0)
                    if (wParam == (IntPtr)WM_KEYDOWN)   //<- (for more performance, if is a keyboard key pressed down event, send callback)
                        if (OnPressKeys != null)
                            OnPressKeys(Marshal.ReadInt32(lParam));

                return CallNextHookEx(hookId, nCode, wParam, lParam);
            }

            public void Dispose()
            {
                //If is not disposed, dispose of this object
                if (isDisposed == false)
                {
                    //Remove all low level hooks
                    UnhookWindowsHookEx(hookId);

                    //Clean variables
                    procedure = null;
                    hookId = IntPtr.Zero;
                }
                isDisposed = true;
            }
        }

        public class KeyboardHotkey_Interceptor : IDisposable
        {
            //Private variables
            private WindowInteropHelper host;
            private int identifier;
            private bool isDisposed = false;

            //Public enums
            [Flags]
            public enum ModifierKeyCodes : uint
            {
                None = 0,
                Alt = 1,
                Control = 2,
                Shift = 4,
                Windows = 8
            }

            //Public variables
            public Window window;
            public ModifierKeyCodes modifier;
            public VirtualKeyInt key;

            //Public callbacks
            public event Action OnPressHotkey;

            //Import methods

            [DllImport("user32.dll")]
            public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

            [DllImport("user32.dll")]
            public static extern bool RegisterHotKey(IntPtr hWnd, int id, ModifierKeyCodes fdModifiers, VirtualKeyInt vk);

            //Core methods

            public KeyboardHotkey_Interceptor(Window window, int id, ModifierKeyCodes modifierCode, VirtualKeyInt keyCode)
            {
                //Store the information
                this.window = window;
                this.modifier = modifierCode;
                this.key = keyCode;

                //Prepare the host and identifier of this hotkey registration
                host = new WindowInteropHelper(this.window);
                identifier = (this.window.GetHashCode() + id);

                //Register the hotkey
                RegisterHotKey(host.Handle, identifier, this.modifier, this.key);

                //Register the callback with a pre-process logic
                ComponentDispatcher.ThreadPreprocessMessage += ProcessMessage;
            }

            void ProcessMessage(ref MSG msg, ref bool handled)
            {
                //Validate the response
                if ((msg.message == 786) && (msg.wParam.ToInt32() == identifier) && (OnPressHotkey != null))
                    OnPressHotkey();
            }

            public void ForceOnPressHotkeyEvent()
            {
                //Force the execution of the event "OnPressHotkey"
                if (OnPressHotkey != null)
                    OnPressHotkey();
            }

            public void Dispose()
            {
                //If is not disposed, dispose of this object
                if (isDisposed == false)
                {
                    //Unregister callback pre-process logic
                    ComponentDispatcher.ThreadPreprocessMessage -= ProcessMessage;

                    //Unregister the hotkey
                    UnregisterHotKey(host.Handle, identifier);
                    this.window = null;
                    host = null;
                }
                isDisposed = true;
            }
        }

        #endregion
    }
}