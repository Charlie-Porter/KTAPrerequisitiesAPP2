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


        //' Initialize global Settings section objects
        private PSCredential psCredentials = null;
        private SqlConnection sqlConnection = null;

        //' Construct dictionaries
        Dictionary<string, string> loadedADKversions = new Dictionary<string, string>();

        public MainWindow()
        {
            InitializeComponent();

            //' Set item source for data grids
            dataGridInstallType.ItemsSource = installTypeCollection;

        }


        private List<string> GetWindowsFeatures(string selection)
        {
            List<string> featureList = new List<string>();

            switch (selection)
            {
                case "TotalAgility WebApp server (Including OPMT)":
                    string[] WebApp = new string[] { "NET-Framework-Features", "NET-WCF-HTTP-Activation" };
                    featureList.AddRange(WebApp);
                    break;
    
            }

            return featureList;
        }

        private void comboBoxInstallType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {

        }

        
        async private void  B_Install_Click(object sender, RoutedEventArgs e)
        {
            //' Clear existing items from observable collection
            if (dataGridInstallType.Items.Count >= 1)
            {
                installTypeCollection.Clear();
            }

            //' Get windows features for selected site type
            List<string> featureList = GetWindowsFeatures(comboBoxInstallType.SelectedItem.ToString());


            //' Update progress bar properties
            progressBarSiteType.Maximum = featureList.Count - 1;
            int progressBarValue = 0;
            labelSiteTypeProgress.Content = string.Empty;

            //' Process each windows feature for installation
            foreach (string feature in featureList)
            {
                //' Update progress bar
                progressBarSiteType.Value = progressBarValue++;
                labelSiteTypeProgress.Content = String.Format("{0} / {1}", progressBarValue, featureList.Count);

                //' Add new item for current windows feature installation state
                installTypeCollection.Add(new WindowsFeature { Name = feature, Progress = true, Result = "Installing..." });
                dataGridInstallType.ScrollIntoView(installTypeCollection[installTypeCollection.Count - 1]);


                //' Invoke windows feature installation via PowerShell runspace
                object installResult = await scriptEngine.AddWindowsFeature(feature);
                string featureState = installResult.ToString();

                //' Update current row on data grid
                if (!String.IsNullOrEmpty(featureState))
                {
                    var currentCollectionItem = installTypeCollection.FirstOrDefault(winFeature => winFeature.Name == feature);

                    if (featureState == "Failed")
                    {
                        if (checkBoxInstallTypeRetryFailed.IsChecked == true)
                        {
                            //' Invoke windows feature installation via PowerShell runspace with alternate source
                            currentCollectionItem.Result = "RetryWithSource";
                            object retryResult = await scriptEngine.AddWindowsFeature(feature, textBoxSettingsSource.Text);
                            featureState = retryResult.ToString();

                            if (featureState == "Failed")
                            {
                                featureState = "FailedAfterRetry";
                            }
                        }
                    }

                    //' Update datagrid elements
                    currentCollectionItem.Progress = false;
                    currentCollectionItem.Result = featureState;
                }

                //' Set color of progressbar
                // new prop needed for binding
            }
        }

        private void button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
