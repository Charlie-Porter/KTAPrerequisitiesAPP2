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
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;

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

        private void Window_Loaded()
        {

            //' Check environment prerequisites
            uint productType = GetProductType();
            switch (productType)
            {
                case 0:
                    MessageBox.Show("Unable to detect platform product type from WMI. Application will now terminate.", "UNHANDLED ERROR", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
                case 1:
                    MessageBox.Show("Unsupported platform detected. Kofax TotalAgility 7.8 supports Windows Server 2008 or above.", "UNSUPPORTED PLATFORM", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
                case 2:
                    MessageBox.Show("Unsupported platform type detect. It's not recommended to run this application on a domain controller.", "WARNING", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;

            }

            bool IsAdmin = IsInGroup(WindowsIdentity.GetCurrent().Token, "Administrators");
            switch (IsAdmin)
            {
                case false:
                    MessageBox.Show("We have detected this app is running as a user who is not a member of the local machine administrators group. If you are a member, please launch this app with the elevated 'Run as Administrator' permission.", "SECURITY ISSUE", MessageBoxButton.OK);
                    Environment.Exit(0);
                    break;
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
        private static string CheckSqlServerUserAccount(string usertobeadded, string password, string server, bool winauth)
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

        private static string TestConnection(string usertobeadded, string password, string server, bool winauth)
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





            // The connection is automatically closed at the end of the using block.
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {


                    connection.Open();


                    return ($@"Connection successful to the SQL server");
                }
                catch (Exception ex)
                {
                    return "Connection failure (you can continue installing by providing a found account if you want to grant the dbcreator role at a different time): " + ex.Message;

                }
            }
        }

        async private void B_Install_Click(object sender, RoutedEventArgs e)
        {
            B_Install.IsEnabled = false;

            dataGridInstallType.Visibility = Visibility.Visible;
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
                if (cb_winauth.IsChecked == true)
                {
                    dbcreatorResult = CheckSqlServerUserAccount(txt_ServiceAcc.Text, txt_sqlpassword.Text, txt_sqlserver.Text, true);

                }
                else
                {
                    dbcreatorResult = CheckSqlServerUserAccount(txt_SQLuser.Text, txt_sqlpassword.Text, txt_sqlserver.Text, false);
                }

                progressBarValue = 1;
                if (!string.IsNullOrEmpty(txt_sqlserver.Text))
                {
                    installTypeCollection.Add(new WindowsFeature { Name = "Grant SQL dbcreator role", Result = dbcreatorResult });
                    dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);
                }

                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 6;

                progressBarValue = 1;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Service Account Logon As A Service rights", Result = GrantUserLogOnAsAService(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 2;
                installTypeCollection.Add(new WindowsFeature { Name = "Install Microsoft Visual C++ Redistributable package", Result = Installvc_redi() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 3;
                installTypeCollection.Add(new WindowsFeature { Name = "Install SQL Server 2012 Native Client", Result = Installsqlncli() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 4;
                installTypeCollection.Add(new WindowsFeature { Name = "Install SQL Server ODBC Driver", Result = Installodbc() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);


                progressBarValue = 5;
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
                progressBarSiteType.Maximum = featureList.Count + 5;

                progressBarValue = 6;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Replace a process level token rights", Result = GrantReplaceAProcessLevelToken(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 7;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Adjust Memory Quotas For A Process", Result = GrantAdjustMemoryQuotasForAProcess(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                progressBarValue = 8;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant Create a token object", Result = GrantCreateATokenObject(txt_ServiceAcc.Text) });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);
            }

            if (comboBoxInstallType.SelectedIndex == 7)
            {
                //' Update progress bar properties
                progressBarSiteType.Maximum = featureList.Count + 5;

                //' Add new item for current windows feature installation state
                progressBarValue = 2;
                installTypeCollection.Add(new WindowsFeature { Name = "Install Microsoft Visual C++ Redistributable package", Result = Installvc_redi() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 3;
                installTypeCollection.Add(new WindowsFeature { Name = "Grant SQL Server 2012 Native Client", Result = Installsqlncli() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);

                //' Add new item for current windows feature installation state
                progressBarValue = 4;
                installTypeCollection.Add(new WindowsFeature { Name = "Install SQL Server ODBC Driver", Result = Installodbc() });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);


                progressBarValue = 5;
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

            B_Install.IsEnabled = true;
            MessageBoxResult result = MessageBox.Show("Prerequisites are applied to this system. You may want to restart the system to ensure all changes are applied. Click Yes to restart or No to cancel.", "System Restart", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
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
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", " /quiet ADDLOCAL=ALL IACCEPTSQLNCLILICENSETERMS=YES /L*V " + Directory.GetCurrentDirectory() + "\\sqlncli.log", path);
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
                    result = ex.Message;
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
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", " /quiet ADDLOCAL=ALL IACCEPTMSSQLCMDLNUTILSLICENSETERMS=YES /L*V " + Directory.GetCurrentDirectory() + "\\MsSqlCmdLnUtils.log", path);
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
                    result = ex.Message;
                    return result;
                }
            }
        }

        private string Installodbc()
        {
            string result;

            if (IsSoftwareInstalled("Microsoft ODBC Driver"))
            {
                result = "NoChangeNeeded";
                return result;
            }
            else
            {
                try
                {
                    string path = Directory.GetCurrentDirectory() + $@"\msodbcsql.msi";

                    using (Process myProcess = new Process())
                    {

                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.FileName = "msiexec";
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", " /quiet IACCEPTMSODBCSQLLICENSETERMS=YES ADDLOCAL=ALL /L*V " + Directory.GetCurrentDirectory() + "\\msodbcsql.log", path);
                        myProcess.Start();
                        myProcess.WaitForExit();
                        if (myProcess.ExitCode != 0)
                        {
                            return result = $@"Failed to install Microsoft ODBC Driver with error code: {myProcess.ExitCode.ToString()}";
                        }
                        return result = "Installed successfully";
                    }
                }
                catch (Exception ex)
                {
                    result = ex.Message;
                    return result;
                }
            }
        }

        private string Installvc_redi()
        {
            string result;

            if (IsSoftwareInstalled("Microsoft Visual C++ 2019 X64"))
            {
                result = "NoChangeNeeded";
                return result;
            }
            else
            {
                try
                {
                    string path = Directory.GetCurrentDirectory() + $@"\VC_redist_2019_x64.exe";

                    using (Process myProcess = new Process())
                    {

                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.FileName = "VC_redist_2019_x64.exe";
                        myProcess.StartInfo.Arguments = (" /q /norestart");
                        myProcess.Start();
                        myProcess.WaitForExit();

                        if (myProcess.ExitCode != 0)
                        {
                            return result = $@"Failed to install Microsoft Visual C++ Redistributable package with error code: {myProcess.ExitCode.ToString()}";
                        }
                        return result = "Installed successfully";
                    }
                }
                catch (Exception ex)
                {
                    result = ex.Message;
                    return result;
                }
            }
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
        
       

        private bool IsInGroup(IntPtr Token, string group)
        {
            using (var identity = new WindowsIdentity(Token))
            {
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(group);
            }
        }
        /// <summary>
        /// https://stephenhaunts.com/2013/03/04/checking-a-user-in-active-directory/
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="username"></param>
        /// <returns></returns>
        public static bool CheckUserinAD(string domain, string username)
        {
            using (var domainContext = new PrincipalContext(ContextType.Domain, domain))
            {
                using (var user = new UserPrincipal(domainContext))
                {
                    user.SamAccountName = username;

                    using (var pS = new PrincipalSearcher())
                    {
                        pS.QueryFilter = user;

                        using (PrincipalSearchResult<Principal> results = pS.FindAll())
                        {
                            if (results != null && results.Count() > 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private uint GetProductType()
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

        private void cb_winauth_Checked(object sender, RoutedEventArgs e)
        {
            if (cb_winauth.IsChecked == true)
            {
                txt_SQLuser.Visibility = Visibility.Collapsed;
                txt_sqlpassword.Visibility = Visibility.Collapsed;
                l_sqlpassword.Visibility = Visibility.Collapsed;
                l_sqluser.Visibility = Visibility.Collapsed;
                txt_dbcreator.Visibility = Visibility.Visible;
                
                
            }
            else
            {
                txt_SQLuser.Visibility = Visibility.Visible;
                txt_sqlpassword.Visibility = Visibility.Visible;
                l_sqlpassword.Visibility = Visibility.Visible;
                l_sqluser.Visibility = Visibility.Visible;
                txt_dbcreator.Visibility = Visibility.Collapsed;
            }
        }

        private void txt_ServiceAcc_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                char[] charSeparators = new char[] { '\\' };
                var ServiceAcc = txt_ServiceAcc.Text.Split(charSeparators);

                if (!CheckUserinAD(ServiceAcc[0], ServiceAcc[1]))
                {

                        

                    if (!localUserExists(txt_ServiceAcc.Text))
                    {
                        tb_message.Text = "Failure - The service account cannot be found in Active Directory or local machine. The install button will be disabled until the account is found.";
                        B_Install.IsEnabled = false;
                    }
                    else
                    {

                        tb_message.Text = "Success - The service account was found in your local machine. ";
                        B_Install.IsEnabled = true;
                    }


                }
                else
                {
                    tb_message.Text = "Success - The service account was found in Active Directory.";
                    B_Install.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (!localUserExists(txt_ServiceAcc.Text))
                    {
                        tb_message.Text = $@"Failure - The service account cannot be found in Active Directory - Error:{ex.Message}.  Also, the account was not found in this local machine. The install button will be disabled until the account is found.";
                        B_Install.IsEnabled = false;
                    }
                    else
                    {

                        tb_message.Text = "Success - The service account was found in your local machine. ";
                        B_Install.IsEnabled = true;
                    }
                }
                catch (Exception exc)
                {
                    tb_message.Text = $@"Failure - The service account cannot be found on this local machine  -Error:{exc.Message}. Also, the account was not found in Active Directory - Error:{ex.Message}. The install button will be disabled until the account is found.";
                    B_Install.IsEnabled = false;
                }
               
    
            }



                
               
        }

        private void txt_dbcreator_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                char[] charSeparators = new char[] { '\\' };
                var dbcreatorAcc = txt_dbcreator.Text.Split(charSeparators);

                if (!CheckUserinAD(dbcreatorAcc[0], dbcreatorAcc[1]))
                {



                    if (!localUserExists(txt_dbcreator.Text))
                    {
                        tb_message.Text = "Failure - The windows account for DB creation cannot be found in Active Directory or local machine. The install button will be disabled until the account is found.";
                        B_Install.IsEnabled = false;
                    }
                    else
                    {

                        tb_message.Text = "Success - The windows account for DB creation was found in your local machine. ";
                        B_Install.IsEnabled = true;
                    }


                }
                else
                {
                    tb_message.Text = "Success - The windows account for DB creation was found in Active Directory.";
                    B_Install.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (!localUserExists(txt_dbcreator.Text))
                    {
                        tb_message.Text = $@"Failure - The windows account for DB creation was found in Active Directory - Error:{ex.Message}.  Also, the account was not found in this local machine. The install button will be disabled until the account is found.";
                        B_Install.IsEnabled = false;
                    }
                    else
                    {

                        tb_message.Text = "Success - The windows account for DB creation was found was found in your local machine. ";
                        B_Install.IsEnabled = true;
                    }
                }
                catch (Exception exc)
                {
                    tb_message.Text = $@"Failure - The windows account for DB creation was found on this local machine  -Error:{exc.Message}. Also, the account was not found in Active Directory - Error:{ex.Message}. The install button will be disabled until the account is found.";
                    B_Install.IsEnabled = false;
                }


            }





        }

        static bool localUserExists(string User)
        {

            using (PrincipalContext pc = new PrincipalContext(ContextType.Machine))
            {
                UserPrincipal up = UserPrincipal.FindByIdentity(
                    pc,
                    IdentityType.SamAccountName,
                    User);
                bool UserExists = (up != null);
                return UserExists;
            }
            
        }

        private void txt_sqlserver_LostFocus(object sender, RoutedEventArgs e)
        {
            
           
        }

        private void b_testconnection_Click(object sender, RoutedEventArgs e)
        { 
            if (cb_dbcreator.IsChecked == true)
            {

                tb_message.Text = TestConnection(txt_dbcreator.Text, txt_sqlpassword.Text, txt_sqlserver.Text, cb_winauth.IsEnabled);
            }
            else
            {
                tb_message.Text = TestConnection(txt_ServiceAcc.Text, txt_sqlpassword.Text, txt_sqlserver.Text, cb_winauth.IsEnabled);
            }

            
        
        }

        private void dataGridInstallType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void cb_dbcreator_Checked(object sender, RoutedEventArgs e)
        {
            if (cb_dbcreator.IsChecked == true)
            {
               
                groupBox.Visibility = Visibility.Visible;

            }
            else
            {
                groupBox.Visibility = Visibility.Collapsed;
                

            }
        }

      
       
    }
}
