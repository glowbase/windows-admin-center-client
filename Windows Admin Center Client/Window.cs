using System;
using System.IO;
using System.Net;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;

using CefSharp;
using DiscordRPC;
using CefSharp.WinForms;

namespace Windows_Admin_Center_Client
{
    public partial class Window : Form
    {
        public string WINDOWS_ADMIN_CENTER_URL = "https://localhost:6516/";
        public string WINDOW_ADMIN_CENTER_INSTALLER = "Installer.msi";
        public string WINDOWS_ADMIN_CENTER_DOWNLOAD = "https://aka.ms/wacdownload/";
        public string DISCORD_CLIENT_ID = "807837695975751690";
        public string CURRENT_DIRECTORY = Environment.CurrentDirectory + @"\";
        public string WINDOWS_ADMIN_CENTER = @"C:\Program Files\Windows Admin Center\SmeDesktop.exe";

        public DiscordRpcClient RPCClient;
        public ChromiumWebBrowser BrowserControl;

        public Dictionary<string, string> Connections;
        public Dictionary<string, string> ServerSettings;
        public Dictionary<string, string> AppSettings;
        public Dictionary<string, string> Tools;

        public Window()
        {
            ComponentResourceManager resources = new ComponentResourceManager(
                typeof(Window)
            );

            this.SuspendLayout();

            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1350, 730);
            this.Icon = ((Icon)(Properties.Resources.Icon));
            this.Name = "Window";

            this.ResumeLayout();

            this.Load += Window_Load;
            this.FormClosing += Window_FormClosing;

            CheckForIllegalCrossThreadCalls = false;

            if (Process.GetProcessesByName("SmeDesktop").Length == 0)
            {
                Process.Start(WINDOWS_ADMIN_CENTER);
            }

            RPCClient = new DiscordRpcClient(DISCORD_CLIENT_ID);
            RPCClient.Initialize();

            Connections = new Dictionary<string, string>()
            {
                {"servermanager", "Server Manager"},
                {"computerManagement", "Computer Management"},
                {"clustermanager", "Cluster Manager"},
            };

            AppSettings = new Dictionary<string, string>()
            {
                {"account", "Account"},
                {"personalization", "Personalisation"},
                {"language", "Language/Region"},
                {"notificaton", "Notifications"},
                {"advanced", "Advanced"},
                {"performance", "Performance Profile"},
                {"extension", "Extensions"},
                {"azure", "Azure"},
                {"connectivity", "Internet Access"},
            };

            ServerSettings = new Dictionary<string, string>()
            {
                {"smbServer", "File Shares (SMB Server)"},
                {"environmentVariables", "Environment Variables"},
                {"hybridManagement", "Azure Arc for Servers"},
                {"powerConfiguration", "Power Configuration"},
                {"remoteDesktop", "Remove Desktop"},
                {"rbacSettings", "Role-based Access Control"},
                {"hyperVHostGeneral", "Hyper-V - General"},
                {"hyperVHostEnhancedSessionMode", "Hyper-V - Enhanced Session Mode"},
                {"hyperVHostNumaSpanning", "Hyper-V - NUMA Spanning"},
                {"hyperVHostLiveMigration", "Hyper-V - Live Migration"},
                {"hyperVHostStorageMigration", "Hyper-V Storage Migration"},
            };

            Tools = new Dictionary<string, string>()
            {
                {"overview", "Overview"},
                {"appsandfeatures", "Apps & Features"},
                {"azure-hybrid-services", "Azure Hybrid Services"},
                {"ad", "Active Directory"},
                {"backup", "Azure Backup"},
                {"azurefilesync", "Azure File Sync"},
                {"azuremonitoring", "Azure Monitor"},
                {"azure-security-center", "Azure Security Center"},
                {"certificates", "Certificates"},
                {"containers", "Containers"},
                {"devices", "Devices"},
                {"dns", "DNS"},
                {"events", "Events"},
                {"files", "Files & File Sharing"},
                {"firewall", "Firewall"},
                {"apps", "Installed Apps"},
                {"usersgroups", "Local Users & Groups"},
                {"network", "Networks"},
                {"monitor", "Performance Monitor"},
                {"powershell", "PowerShell"},
                {"processes", "Processes"},
                {"registry", "Registry"},
                {"remotedesktop", "Remote Desktop"},
                {"rolesfeatures", "Roles & Features"},
                {"scheduledtasks", "Scheduled Tasks"},
                {"security", "Security"},
                {"services", "Services"},
                {"storage", "Storage"},
                {"storage-migration", "Storage Migration Service"},
                {"storage-replica", "Storage Replica"},
                {"system-insights", "System Insights"},
                {"windowsupdate", "Windows Updates"},
                {"virtualmachines", "Virtual Machines"},
                {"virtualswitches", "Virtual Switches"},
            };
        }

        private void Window_Load(object sender, EventArgs e)
        {
            if (!File.Exists(WINDOWS_ADMIN_CENTER))
            {
                if (!File.Exists(CURRENT_DIRECTORY + WINDOW_ADMIN_CENTER_INSTALLER))
                {
                    using (WebClient Client = new WebClient())
                    {
                        Client.DownloadFile(WINDOWS_ADMIN_CENTER_DOWNLOAD, WINDOW_ADMIN_CENTER_INSTALLER);
                    }
                }

                Process.Start(CURRENT_DIRECTORY + WINDOW_ADMIN_CENTER_INSTALLER);
                Application.Exit();
            }

            BrowserControl = new ChromiumWebBrowser(
                WINDOWS_ADMIN_CENTER_URL
            );

            CefSettings BrowserSettings = new CefSettings
            {
                CachePath = CURRENT_DIRECTORY + "Cache",
                UserDataPath = CURRENT_DIRECTORY + "UserData"
            };

            if (!Cef.IsInitialized)
            {
                Cef.Initialize(BrowserSettings);
            }

            this.Controls.Add(BrowserControl);
            BrowserControl.Dock = DockStyle.Fill;

            BrowserControl.TitleChanged += Browser_TitleChanged;
        }

        private void Browser_TitleChanged(object sender, TitleChangedEventArgs e)
        {
            string Title = e.Title;
            string URL = BrowserControl.Address.Replace(WINDOWS_ADMIN_CENTER_URL, "");
            string[] Breakdown = URL.Split("/".ToCharArray());

            string MachineType = null;
            string ToolType = null;
            string HostName = null;

            string RPCDetails = null;
            string RPCState = null;

            string RPCSmallIcon = null;
            string RPCSmallIconText = null;
            string RPCLargeIcon = null;
            string RPCLargeIconText = null;

            this.Text = Title;


            if (URL.EndsWith("connections"))
            {
                RPCDetails = Connections[Breakdown[0]];
                RPCState = "In Menu";

                RPCLargeIcon = "wac";
                RPCLargeIconText = "Windows Admin Center";
            }
            else if (Breakdown.Length == 1)
            {
                RPCDetails = "All Connections";
                RPCState = "In Menu";

                RPCLargeIcon = "wac";
                RPCLargeIconText = "Windows Admin Center";
            } else if (URL.Contains("tools"))
            {
                MachineType = Breakdown[2];
                HostName = Breakdown[3];
                ToolType = Breakdown[5].ToLower();

                if (ToolType == "settings")
                {
                    string SettingType = Breakdown[6];

                    RPCDetails = HostName;
                    RPCState = "Settings: " + ServerSettings[SettingType];
                    RPCSmallIconText = ServerSettings[SettingType];
                    RPCLargeIcon = "wac";
                    RPCLargeIconText = "Windows Admin Center";
                } else
                {
                    RPCDetails = HostName;
                    RPCState = Tools[ToolType];
                    RPCLargeIcon = ToolType;
                    RPCLargeIconText = Tools[ToolType];
                    RPCSmallIcon = "wac";
                    RPCSmallIconText = "Windows Admin Center";
                }
            } else if (URL.StartsWith("settings"))
            {
                string SettingType = Breakdown[1];
                string TempDetails = null;

                if (SettingType == "extension")
                {
                    string ExtensionType = Breakdown[2];

                    TempDetails = AppSettings[SettingType] + 
                        " (" + 
                            ExtensionType[0].ToString().ToUpper() + 
                            ExtensionType.Substring(1) + 
                        ")";
                } else
                {
                    TempDetails = AppSettings[SettingType];
                }

                RPCDetails = TempDetails;
                RPCState = "Settings Menu";
                RPCSmallIconText = SettingType;
                RPCLargeIcon = "wac";
                RPCLargeIconText = "Windows Admin Center";
            }

            RPCClient.SetPresence(new RichPresence()
            {
                Details = RPCDetails,
                State = RPCState,
                Timestamps = Timestamps.Now,
                Assets = new Assets()
                {
                    LargeImageKey = RPCLargeIcon,
                    LargeImageText = RPCLargeIconText,
                    SmallImageKey = RPCSmallIcon,
                    SmallImageText = RPCSmallIconText
                }
            });
        }

        private void Window_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}