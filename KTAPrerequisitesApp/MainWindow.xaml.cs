using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Management;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Windows.Media.Imaging;

namespace KTAPrerequisitesApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml 
    /// https://github.com/AngryCarrot789/WPFDarkTheme
    /// https://www.youtube.com/watch?v=5OdkVXW5Z0E
    /// </summary>
    public partial class MainWindow : Window
    {
        //' Construct new class objects
        ScriptEngine scriptEngine = new ScriptEngine();

        //' Construct observable collections as datagrids item source
        private ObservableCollection<WindowsFeature> installTypeCollection = new ObservableCollection<WindowsFeature>();


      
        public MainWindow()
        {
            Window_Loaded();
            InitializeComponent();
            Uri iconUri = new Uri("pack://application:,,,/Resources/kofaxlogo.png", UriKind.RelativeOrAbsolute);
            this.Icon = BitmapFrame.Create(iconUri);
        }


        private List<string> GetWindowsFeatures(string selection)
        {
            List<string> featureList = new List<string>();

            switch (selection)
            {
                case "TotalAgility WebApp server(Including OPMT)":
                    string[] WebApp = new string[] { "Web-Mgmt-Console", "NET-Framework-Features",  "Web-Windows-Auth","Web-Asp-Net45","Web-Static-Content","NET-Framework-Core","NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(WebApp);
                    break;
                    case "TotalAgility Web Only(Including OPMT)":
                    string[] Web = new string[] { "Web-Mgmt-Console", "NET-Framework-Features",  "Web-Windows-Auth","Web-Asp-Net45","Web-Static-Content","NET-Framework-Core","NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(Web);
                    break;
                case "TotalAgility APP Only(Including OPMT)":
                    string[] APP = new string[] { "Web-Mgmt-Console", "NET-Framework-Features", "Web-Windows-Auth", "Web-Asp-Net45", "Web-Static-Content", "NET-Framework-Core", "NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(APP);
                    break;
                case "TotalAgility Transformation Server":
                    string[] TS = new string[] { "" };
                    featureList.AddRange(TS);
                    break;
                    case "TotalAgility Transformation Server(OPMT)":
                    string[] TSOPMT = new string[] { "", };
                    featureList.AddRange(TSOPMT);
                    break;
                    case "TotalAgility Intergration Server":
                    string[] IS = new string[] { "Web-Mgmt-Console","NET-Framework-Features",  "Web-Windows-Auth","Web-Asp-Net45","Web-Static-Content","NET-Framework-Core","NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(IS);
                    break;
                    case "TotalAgility RTTS":
                    string[] RTTS = new string[] { "Web-Mgmt-Console", "NET-Framework-Features", "Web-Windows-Auth","Web-Asp-Net45","Web-Static-Content","NET-Framework-Core","NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(RTTS);
                    break;
                    case "TotalAgility DB Only":
                    string[] DB = new string[] { "", };
                    featureList.AddRange(DB);
                    break;
            }

            return featureList;
        }

        private string Installsqlncli()
        {
            string result;

            if (IsSoftwareInstalled("Microsoft SQL Server 2012 Native Client "))
            {
                result = "NoChangeNeeded";
                return result;
            }
            else
            {
                try
                {

                    string path = Directory.GetCurrentDirectory() + $@"\sqlncli.msi";
                   
                    using (Process myProcess = new Process())
                    {

                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.FileName = "msiexec";
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", "ADDLOCAL=ALL /quiet IACCEPTSQLNCLILICENSETERMS=YES ADDLOCAL=ALL /L*V " + Directory.GetCurrentDirectory() + "\\sqlncli.log", path);
                        myProcess.Start();
                        myProcess.WaitForExit();
                        if (myProcess.ExitCode != 0)
                        {
                            return result = $@"Failed to install SQL Native Client tools with error code: {myProcess.ExitCode.ToString()}";
                        }
                        return result = "Installed successfully";
                    }
                }
                catch (Exception ex)
                {
                    result = ex.ToString();
                    return result;
                }
            }
        }


        private string Installsqlcmd()
        {
            string result;

            if (IsSoftwareInstalled("Microsoft Command Line Utilities"))
            {
                result = "NoChangeNeeded";
                return result;
            }
            else
            {
                try
                {
                    string path = Directory.GetCurrentDirectory() + $@"\MsSqlCmdLnUtils.msi";

                    using (Process myProcess = new Process())
                    {

                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.FileName = "msiexec";
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", "ADDLOCAL=ALL /quiet IACCEPTSQLNCLILICENSETERMS=YES ADDLOCAL=ALL /L*V " + Directory.GetCurrentDirectory() + "\\sqlncli.log", path);
                        myProcess.Start();
                        myProcess.WaitForExit();
                        if (myProcess.ExitCode != 0)
                        {
                            return result = $@"Failed to install SQL Command Line Utilitys with error code: {myProcess.ExitCode.ToString()}";
                        }
                        return result = "Installed successfully";
                    }
                }
                catch (Exception ex)
                {
                    result = ex.ToString();
                    return result;
                }
            }
        }
        /// <summary>
        /// https://stackoverflow.com/questions/56548873/c-sharp-sql-server-add-roles-to-database-user
        /// </summary>
        /// <param name="usertobeadded"></param>
        /// <param name="password"></param>
        /// <param name="server"></param>
        /// <param name="winauth"></param>
        /// <returns></returns>
        private static string CheckSqlServerUserAccount(string usertobeadded, string password, string server, bool winauth )
        {
            string connectionString;

            if (winauth == false)
            {
                connectionString = $@"Data Source={server};User ID={usertobeadded};Password={password}";
            }
            else
            {
                connectionString = $@"Data Source={server};Trusted_Connection=True";
            }

            

            string cmdText = $@"EXEC sp_addsrvrolemember '{usertobeadded}', 'dbcreator';";

            // The connection is automatically closed at the end of the using block.
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand(cmdText, connection);
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    return ($@"Granted DBCreator role to {usertobeadded}");
                }
                catch (Exception ex)
                {
                    return ex.Message;
                    
                }
            }
        }


        async private void B_Install_Click(object sender, RoutedEventArgs e)
        {

            //' Clear existing items from observable collection
            if (dataGridInstallType.Items.Count >= 1)
            {
                installTypeCollection.Clear();
            }

            //' Set item source for data grids
            dataGridInstallType.ItemsSource = installTypeCollection;


            //' Get windows features for selected site type
            List<string> featureList = GetWindowsFeatures(comboBoxInstallType.SelectedItem.ToString());

            int progressBarValue = 0;

            //0 TotalAgility WebApp server(Including OPMT)

            //  1  TotalAgility Web Only(Including OPMT)

            //    2  TotalAgility APP Only(Including OPMT)

            //     3   TotalAgility Transformation Server

            //        4    TotalAgility Transformation Server(OPMT)

            //         5     TotalAgility Intergration Server

            //          6   TotalAgility RTTS 

            //            7    TotalAgility DB Only


            if (comboBoxInstallType.SelectedIndex == 0 || comboBoxInstallType.SelectedIndex == 2 || comboBoxInstallType.SelectedIndex == 3 || comboBoxInstallType.SelectedIndex == 4 || comboBoxInstallType.SelectedIndex == 5 || comboBoxInstallType.SelectedIndex == 6) 
            {

                
                string dbcreatorResult;
                if (cb_WinAuth.IsChecked == true)
                {
                    dbcreatorResult = CheckSqlServerUserAccount(txt_ServiceAcc.Text, txt_sqlpassword.Text, txt_sqlserver.Text, true);

                }
                else
                {
                    dbcreatorResult = CheckSqlServerUserAccount(txt_SQLuser.Text, txt_sqlpassword.Text, txt_sqlserver.Text, false);
                }

                progressBarValue = 1;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant SQL dbcreator role", Result = dbcreatorResult });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);


                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 4;
                
                progressBarValue = 1;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Service Account Logon As A Service rights", Result = GrantUserLogOnAsAService(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 2;
                installTypeCollection.Add(new WindowsFeature { Name = "Install SQL Server 2012 Native Client", Result = Installsqlncli() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 3;
                installTypeCollection.Add(new WindowsFeature { Name = "Install SQL Server Command Line Utility", Result = Installsqlcmd() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

            }
            if (comboBoxInstallType.SelectedIndex == 1) // Web Only Server OPMT
            {
                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 2;
                progressBarValue = 1;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Service Account Logon As A Service rights", Result = GrantUserLogOnAsAService(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

            }

            if (comboBoxInstallType.SelectedIndex == 4) // Transformation Sercer OPMT
            {

                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 4;
                progressBarValue = 4;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Replace a process level token rights", Result = GrantReplaceAProcessLevelToken(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 5;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Adjust Memory Quotas For A Process", Result = GrantAdjustMemoryQuotasForAProcess(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 6;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Create a token object", Result = GrantCreateATokenObject(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);
            }

            if (comboBoxInstallType.SelectedIndex == 7)
            {
                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 3;
                //' Add new item for current windows feature installation state
                progressBarValue = 2;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant SQL Server 2012 Native Client", Result = Installsqlncli() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 3;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant SQL Server Command Line Utility", Result = Installsqlcmd() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);
            }


            //' Process each windows feature for installation
            foreach (string feature in featureList)
            {
                if (!String.IsNullOrEmpty(feature))
                {
                    //' Update progress bar
                    progressBarSiteType.Value = progressBarValue++;
                    

                    //' Add new item for current windows feature installation state
                    installTypeCollection.Add(new WindowsFeature { Name = feature, Result = "Granting..." });
                    dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);


                    //' Invoke windows feature installation via PowerShell runspace
                    object installResult = await scriptEngine.AddWindowsFeature(feature);
                    string featureState = installResult.ToString();

                    //' Update current row on data grid
                    if (!String.IsNullOrEmpty(featureState))
                    {
                        var currentCollectionItem = installTypeCollection.FirstOrDefault(winFeature => winFeature.Name == feature);

                        //' Update datagrid elements
                        currentCollectionItem.Result = featureState;
                    }
                }

            }
            progressBarSiteType.Value = progressBarSiteType.Maximum;

            MessageBoxResult result = MessageBox.Show("Prerequisites are applied to this system. You may want to restart the system to ensure all changes are applied. Click Yes to restart or No to cancel.", "System Restart", MessageBoxButton.YesNo);
            
            if (result== MessageBoxResult.Yes)
            {
                SystemRestart();
            } 
        }

        private void SystemRestart()
        {
            var cmd = new ProcessStartInfo("shutdown.exe", "-r -t 0");
            cmd.CreateNoWindow = true;
            cmd.UseShellExecute = false;
            cmd.ErrorDialog = false;
            Process.Start(cmd);
        }



        private static string GrantUserLogOnAsAService(string userName)
        {
            try
            {
                LsaWrapper lsaUtility = new LsaWrapper();
                lsaUtility.SetRight(userName, "SeServiceLogonRight");
                return ("Granted Logon as a Service right");
            }
            catch (Exception ex)
            {
                return(ex.Message);
            }
        }

        private static string GrantReplaceAProcessLevelToken(string userName)
        {
            try
            {
                LsaWrapper lsaUtility = new LsaWrapper();
                lsaUtility.SetRight(userName, "SeAssignPrimaryTokenPrivilege");
                return ("Granted Replace a process level token right");
            }
            catch (Exception ex)
            {
                return (ex.Message);
            }
        }

        private static string GrantCreateATokenObject(string userName)
        {
            try
            {
                LsaWrapper lsaUtility = new LsaWrapper();
                lsaUtility.SetRight(userName, "SeCreateTokenPrivilege");
                return ("Granted Create a token object right");
            }
            catch (Exception ex)
            {
                return (ex.Message);
            }
        }


        public static string GrantAdjustMemoryQuotasForAProcess(string userName)
        {
            try
            {
                LsaWrapper lsaUtility = new LsaWrapper();
                lsaUtility.SetRight(userName, "SeIncreaseQuotaPrivilege");
                return ("Granted Adjust memory quotas for a process");
            }
            catch (Exception ex)
            {
                return (ex.Message);
            }
        }

        private void comboBoxInstallType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        /// <summary>
        /// https://stackoverflow.com/questions/16379143/check-if-application-is-installed-in-registry
        /// </summary>
        /// <param name="softwareName"></param>
        /// <returns></returns>
        private static bool IsSoftwareInstalled(string softwareName)
        {
            var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") ??
                      Registry.LocalMachine.OpenSubKey(
                          @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall");

            if (key == null)
                return false;

            return key.GetSubKeyNames()
                .Select(keyName => key.OpenSubKey(keyName))
                .Select(subkey => subkey.GetValue("DisplayName") as string)
                .Any(displayName => displayName != null && displayName.Contains(softwareName));
        }
        
        private void Window_Loaded()
        {
            MessageBoxResult result;
            //' Check environment prerequisites
            uint productType = GetProductType();
            switch (productType)
            {
                case 0:
                    result = MessageBox.Show("Unable to detect platform product type from WMI. Application will now terminate.", "UNHANDLED ERROR", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
                case 1:
                    result = MessageBox.Show("Unsupported platform detected. Kofax TotalAgility supports Windows 2008 or above.", "UNSUPPORTED PLATFORM", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
                case 2:
                    result = MessageBox.Show("Unsupported platform type detect. It's not recommended to run this application on a domain controller.", "WARNING", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
                
            }
            



        }

       

    public uint GetProductType()
        {
            uint productType = 0;

            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject managementObject in searcher.Get())
                {
                    productType = (uint)managementObject.GetPropertyValue("ProductType");
                }
            }

            return productType;
        }

        private void cb_WinAuth_Checked(object sender, RoutedEventArgs e)
        {
            if (cb_WinAuth.IsChecked == true)
            {
                txt_SQLuser.IsEnabled = false;
                txt_sqlpassword.IsEnabled = false;
                l_sqluser.IsEnabled = false;
                l_sqlpassword.IsEnabled = false;
            }
            else
            {
                txt_SQLuser.IsEnabled = true;
                txt_sqlpassword.IsEnabled = true;
                l_sqluser.IsEnabled = true;
                l_sqlpassword.IsEnabled = true;
            }
        }
    }
}
