using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Collections;
using System.Data;
using System.Net;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows.Data;
using System.Text;
using System.Windows.Markup;
using System.Globalization;
using System.Collections.Specialized;
using System.Net.NetworkInformation;
using System.Management;
using WindowsInstaller;
using Microsoft.Win32;

namespace KTAPrerequisitesApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //' Construct new class objects
        ScriptEngine scriptEngine = new ScriptEngine();

        //' Construct observable collections as datagrids item source
        private ObservableCollection<WindowsFeature> installTypeCollection = new ObservableCollection<WindowsFeature>();


      
        public MainWindow()
        {
            InitializeComponent();
        }


        private List<string> GetWindowsFeatures(string selection)
        {
            List<string> featureList = new List<string>();

            switch (selection)
            {
                case "TotalAgility WebApp server (Including OPMT)":
                    string[] WebApp = new string[] { "NET-Framework-Features", "NET-HTTP-Activation", "Web-Windows-Auth","Web-Asp-Net45","Web-Static-Content","NET-Framework-Core","NET-WCF-HTTP-Activation45", };
                    featureList.AddRange(WebApp);
                    break;
    
            }

            return featureList;
        }

        private string Installmsi()
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
                    //Type type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                    //Installer installer = (Installer)Activator.CreateInstance(type);
                    //installer.UILevel = MsiUILevel.msiUILevelNone;

                    string path = Directory.GetCurrentDirectory() + $@"\sqlncli.msi";
                    //installer.InstallProduct($@"{directory}\sqlncli.msi", "ACTION=ADMIN");

                    //return result = "Installed";
                    using (Process myProcess = new Process())
                    {

                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.FileName = "msiexec";
                        myProcess.StartInfo.Arguments = string.Format(" /i {1} {0}", "ADDLOCAL=ALL /quiet IACCEPTSQLNCLILICENSETERMS=YES ADDLOCAL=ALL /L*V " + Directory.GetCurrentDirectory() + "\\sqlncli.log", path);
                        myProcess.Start();
                        myProcess.WaitForExit();
                        return result = "Installed";
                    }
                }
                catch (Exception ex)
                {
                    result = ex.ToString();
                    return result;
                }
            }
        }


        
        async private void  B_Install_Click(object sender, RoutedEventArgs e)
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


            //' Update progress bar properties
            progressBarSiteType.Maximum = featureList.Count - 1;
            int progressBarValue = 0;
            labelSiteTypeProgress.Content = string.Empty;

            //' Add new item for current windows feature installation state
            progressBarValue = 1;                     
            string result = Installmsi();
            installTypeCollection.Add(new WindowsFeature { Name = "SQL Server 2012 Native Client", Result = result });
            dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);
            //' Process each windows feature for installation
            foreach (string feature in featureList)
            {
                //' Update progress bar
                progressBarSiteType.Value = progressBarValue++;
                labelSiteTypeProgress.Content = String.Format("{0} / {1}", progressBarValue, featureList.Count + 1);

                //' Add new item for current windows feature installation state
                installTypeCollection.Add(new WindowsFeature { Name = feature, Result = "Installing..." });
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
    }
}
